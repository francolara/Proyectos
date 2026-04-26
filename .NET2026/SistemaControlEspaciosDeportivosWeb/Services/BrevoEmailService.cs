using System.Net.Http.Json;
using Microsoft.Extensions.Options;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class BrevoEmailService(
    HttpClient httpClient,
    IOptions<BrevoSettings> options,
    ILogger<BrevoEmailService> logger) : IEmailService
{
    private const string EndpointCorreoTransaccional = "smtp/email";
    private static readonly HashSet<string> ContentTypesPermitidosPorDefecto = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/pdf",
        "image/jpeg",
        "image/png",
        "text/plain"
    };
    private readonly BrevoSettings _settings = options.Value;

    public async Task SendEmailAsync(
        string toEmail,
        string toName,
        string subject,
        string htmlContent,
        EmailSendOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(_settings.ApiKey) ||
            string.IsNullOrWhiteSpace(_settings.SenderEmail) ||
            string.IsNullOrWhiteSpace(_settings.SenderName))
        {
            throw new EmailDeliveryException("La configuracion de Brevo no esta completa.");
        }

        if (string.IsNullOrWhiteSpace(toEmail))
        {
            throw new EmailDeliveryException("El correo de destino es obligatorio.");
        }

        if (string.IsNullOrWhiteSpace(subject))
        {
            throw new EmailDeliveryException("El asunto es obligatorio.");
        }

        if (string.IsNullOrWhiteSpace(htmlContent))
        {
            throw new EmailDeliveryException("El contenido HTML es obligatorio.");
        }

        var senderEmail = ResolverSenderEmail(options?.SenderEmail);
        var senderName = string.IsNullOrWhiteSpace(options?.SenderName) ? _settings.SenderName : options!.SenderName!.Trim();
        var attachments = await ConstruirAdjuntosAsync(options?.AttachmentUrls);

        var payload = new Dictionary<string, object?>
        {
            ["sender"] = new
            {
                name = senderName,
                email = senderEmail
            },
            ["to"] = new[]
            {
                new
                {
                    email = toEmail,
                    name = string.IsNullOrWhiteSpace(toName) ? toEmail : toName
                }
            },
            ["subject"] = subject,
            ["htmlContent"] = htmlContent
        };

        if (attachments.Count > 0)
        {
            payload["attachment"] = attachments;
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, EndpointCorreoTransaccional);
        request.Headers.TryAddWithoutValidation("api-key", _settings.ApiKey);
        request.Content = JsonContent.Create(payload);

        try
        {
            using var response = await httpClient.SendAsync(request);
            if (response.IsSuccessStatusCode) return;

            var detalle = await response.Content.ReadAsStringAsync();
            logger.LogError(
                "Brevo devolvio error al enviar correo. StatusCode={StatusCode}, Response={Response}.",
                (int)response.StatusCode,
                Truncar(detalle, 500));

            throw new EmailDeliveryException("No se pudo enviar el correo transaccional.");
        }
        catch (EmailDeliveryException)
        {
            throw;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error inesperado al enviar correo transaccional con Brevo.");
            throw new EmailDeliveryException("No se pudo enviar el correo transaccional.", ex);
        }
    }

    private static string Truncar(string? value, int maxLen)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return value.Length <= maxLen ? value : value[..maxLen];
    }

    private string ResolverSenderEmail(string? senderEmailSolicitado)
    {
        if (string.IsNullOrWhiteSpace(senderEmailSolicitado))
        {
            return _settings.SenderEmail.Trim();
        }

        var senderNormalizado = senderEmailSolicitado.Trim();
        if (!_settings.AllowedSenderEmails.Any())
        {
            if (string.Equals(senderNormalizado, _settings.SenderEmail, StringComparison.OrdinalIgnoreCase))
            {
                return senderNormalizado;
            }

            throw new EmailDeliveryException("El remitente solicitado no esta permitido.");
        }

        var permitido = _settings.AllowedSenderEmails.Any(x =>
            string.Equals(x?.Trim(), senderNormalizado, StringComparison.OrdinalIgnoreCase));
        if (!permitido)
        {
            throw new EmailDeliveryException("El remitente solicitado no esta permitido.");
        }

        return senderNormalizado;
    }

    private async Task<List<object>> ConstruirAdjuntosAsync(IEnumerable<EmailAttachmentUrlOption>? attachmentUrls)
    {
        var resultado = new List<object>();
        if (attachmentUrls is null) return resultado;

        foreach (var adjunto in attachmentUrls)
        {
            if (adjunto is null || string.IsNullOrWhiteSpace(adjunto.Url)) continue;
            var item = await DescargarAdjuntoAsync(adjunto);
            resultado.Add(new { name = item.Name, content = item.ContentBase64 });
        }

        return resultado;
    }

    private async Task<AttachmentData> DescargarAdjuntoAsync(EmailAttachmentUrlOption option)
    {
        if (!Uri.TryCreate(option.Url, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            throw new EmailDeliveryException($"La URL del adjunto no es valida: {option.Url}");
        }

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(Math.Max(5, _settings.AttachmentDownloadTimeoutSeconds)));
        using var request = new HttpRequestMessage(HttpMethod.Get, uri);
        using var response = await httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeoutCts.Token);
        if (!response.IsSuccessStatusCode)
        {
            throw new EmailDeliveryException($"No se pudo descargar el adjunto desde {uri}.");
        }

        var limite = Math.Max(1024, _settings.AttachmentMaxBytes);
        var contentLength = response.Content.Headers.ContentLength;
        if (contentLength.HasValue && contentLength.Value > limite)
        {
            throw new EmailDeliveryException($"El adjunto supera el limite permitido de {limite} bytes.");
        }

        var contentType = response.Content.Headers.ContentType?.MediaType;
        ValidarContentTypeAdjunto(contentType);

        await using var stream = await response.Content.ReadAsStreamAsync(timeoutCts.Token);
        using var ms = new MemoryStream();
        await CopiarConLimiteAsync(stream, ms, limite, timeoutCts.Token);

        var fileName = ResolverFileName(option, response, uri);
        return new AttachmentData(fileName, Convert.ToBase64String(ms.ToArray()));
    }

    private void ValidarContentTypeAdjunto(string? contentType)
    {
        if (string.IsNullOrWhiteSpace(contentType))
        {
            throw new EmailDeliveryException("No se pudo validar el tipo de contenido del adjunto.");
        }

        var permitidos = _settings.AllowedAttachmentContentTypes.Any()
            ? _settings.AllowedAttachmentContentTypes
            : ContentTypesPermitidosPorDefecto.ToList();

        var esPermitido = permitidos.Any(x => string.Equals(x?.Trim(), contentType.Trim(), StringComparison.OrdinalIgnoreCase));
        if (!esPermitido)
        {
            throw new EmailDeliveryException($"Tipo de adjunto no permitido: {contentType}");
        }
    }

    private static async Task CopiarConLimiteAsync(Stream input, Stream output, int maxBytes, CancellationToken ct)
    {
        var buffer = new byte[81920];
        var total = 0;
        int leidos;
        while ((leidos = await input.ReadAsync(buffer.AsMemory(0, buffer.Length), ct)) > 0)
        {
            total += leidos;
            if (total > maxBytes)
            {
                throw new EmailDeliveryException($"El adjunto supera el limite permitido de {maxBytes} bytes.");
            }

            await output.WriteAsync(buffer.AsMemory(0, leidos), ct);
        }
    }

    private static string ResolverFileName(EmailAttachmentUrlOption option, HttpResponseMessage response, Uri uri)
    {
        if (!string.IsNullOrWhiteSpace(option.FileName))
        {
            return option.FileName.Trim();
        }

        var fromDisposition = response.Content.Headers.ContentDisposition?.FileNameStar
            ?? response.Content.Headers.ContentDisposition?.FileName;
        if (!string.IsNullOrWhiteSpace(fromDisposition))
        {
            return fromDisposition.Trim().Trim('"');
        }

        var urlName = Path.GetFileName(uri.LocalPath);
        if (!string.IsNullOrWhiteSpace(urlName))
        {
            return urlName;
        }

        return $"adjunto-{Guid.NewGuid():N}.bin";
    }

    private sealed record AttachmentData(string Name, string ContentBase64);
}
