using System.Net.Http.Json;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Email;

public sealed class BrevoEmailService(
    HttpClient httpClient,
    IOptions<BrevoSettings> options,
    ILogger<BrevoEmailService> logger) : IEmailService
{
    private const string TransactionalEmailEndpoint = "smtp/email";
    private static readonly HashSet<string> DefaultAllowedContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/pdf",
        "application/xml",
        "text/xml",
        "application/zip",
        "application/octet-stream",
        "image/jpeg",
        "image/png",
        "text/plain"
    };

    private readonly BrevoSettings _settings = options.Value;

    public bool IsEnabled =>
        !string.IsNullOrWhiteSpace(_settings.ApiKey)
        && !string.IsNullOrWhiteSpace(_settings.SenderEmail)
        && !string.IsNullOrWhiteSpace(_settings.SenderName);

    public async Task SendEmailAsync(
        string toEmail,
        string toName,
        string subject,
        string htmlContent,
        EmailSendOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (!IsEnabled)
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

        var senderEmail = ResolveSenderEmail(options?.SenderEmail);
        var senderName = string.IsNullOrWhiteSpace(options?.SenderName)
            ? _settings.SenderName.Trim()
            : options.SenderName.Trim();
        var attachments = await BuildAttachmentsAsync(options?.AttachmentUrls, cancellationToken);

        var payload = new Dictionary<string, object?>
        {
            ["sender"] = new { name = senderName, email = senderEmail },
            ["to"] = new[]
            {
                new
                {
                    email = toEmail.Trim(),
                    name = string.IsNullOrWhiteSpace(toName) ? toEmail.Trim() : toName.Trim()
                }
            },
            ["subject"] = subject.Trim(),
            ["htmlContent"] = htmlContent
        };

        if (attachments.Count > 0)
        {
            payload["attachment"] = attachments;
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, TransactionalEmailEndpoint);
        request.Headers.TryAddWithoutValidation("api-key", _settings.ApiKey.Trim());
        request.Content = JsonContent.Create(payload);

        try
        {
            using var response = await httpClient.SendAsync(request, cancellationToken);
            if (response.IsSuccessStatusCode)
            {
                return;
            }

            var detail = await response.Content.ReadAsStringAsync(cancellationToken);
            logger.LogError(
                "Brevo rechazo el envio de correo. StatusCode={StatusCode}, Response={Response}.",
                (int)response.StatusCode,
                Truncate(detail, 500));

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

    private string ResolveSenderEmail(string? requestedSenderEmail)
    {
        if (string.IsNullOrWhiteSpace(requestedSenderEmail))
        {
            return _settings.SenderEmail.Trim();
        }

        var normalizedSender = requestedSenderEmail.Trim();
        var allowedSenders = _settings.AllowedSenderEmails
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .ToArray();

        if (allowedSenders.Length == 0)
        {
            if (string.Equals(normalizedSender, _settings.SenderEmail.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return normalizedSender;
            }

            throw new EmailDeliveryException("El remitente solicitado no esta permitido.");
        }

        if (!allowedSenders.Contains(normalizedSender, StringComparer.OrdinalIgnoreCase))
        {
            throw new EmailDeliveryException("El remitente solicitado no esta permitido.");
        }

        return normalizedSender;
    }

    private async Task<List<object>> BuildAttachmentsAsync(
        IEnumerable<EmailAttachmentUrlOption>? attachmentUrls,
        CancellationToken cancellationToken)
    {
        var result = new List<object>();
        if (attachmentUrls is null)
        {
            return result;
        }

        foreach (var attachment in attachmentUrls)
        {
            if (attachment is null || string.IsNullOrWhiteSpace(attachment.Url))
            {
                continue;
            }

            var item = await DownloadAttachmentAsync(attachment, cancellationToken);
            result.Add(new { name = item.Name, content = item.ContentBase64 });
        }

        return result;
    }

    private async Task<AttachmentData> DownloadAttachmentAsync(
        EmailAttachmentUrlOption option,
        CancellationToken cancellationToken)
    {
        if (!Uri.TryCreate(option.Url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            throw new EmailDeliveryException("La URL del adjunto no es valida.");
        }

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(Math.Max(5, _settings.AttachmentDownloadTimeoutSeconds)));
        using var request = new HttpRequestMessage(HttpMethod.Get, uri);
        using var response = await httpClient.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            timeoutCts.Token);

        if (!response.IsSuccessStatusCode)
        {
            throw new EmailDeliveryException("No se pudo descargar uno de los adjuntos.");
        }

        var maxBytes = Math.Max(1024, _settings.AttachmentMaxBytes);
        if (response.Content.Headers.ContentLength is long contentLength && contentLength > maxBytes)
        {
            throw new EmailDeliveryException($"El adjunto supera el limite permitido de {maxBytes} bytes.");
        }

        ValidateAttachmentContentType(response.Content.Headers.ContentType?.MediaType);

        await using var input = await response.Content.ReadAsStreamAsync(timeoutCts.Token);
        using var output = new MemoryStream();
        await CopyWithLimitAsync(input, output, maxBytes, timeoutCts.Token);

        return new AttachmentData(
            ResolveFileName(option, response, uri),
            Convert.ToBase64String(output.ToArray()));
    }

    private void ValidateAttachmentContentType(string? contentType)
    {
        var normalizedContentType = contentType?.Split(';', 2, StringSplitOptions.TrimEntries)[0].Trim();
        if (string.IsNullOrWhiteSpace(normalizedContentType))
        {
            throw new EmailDeliveryException("No se pudo validar el tipo de contenido del adjunto.");
        }

        var allowedContentTypes = _settings.AllowedAttachmentContentTypes.Count > 0
            ? _settings.AllowedAttachmentContentTypes
            : DefaultAllowedContentTypes.ToList();

        if (!allowedContentTypes.Any(x =>
                string.Equals(x?.Trim(), normalizedContentType, StringComparison.OrdinalIgnoreCase)))
        {
            throw new EmailDeliveryException($"Tipo de adjunto no permitido: {normalizedContentType}.");
        }
    }

    private static async Task CopyWithLimitAsync(
        Stream input,
        Stream output,
        int maxBytes,
        CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        var total = 0;
        int read;

        while ((read = await input.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken)) > 0)
        {
            total += read;
            if (total > maxBytes)
            {
                throw new EmailDeliveryException($"El adjunto supera el limite permitido de {maxBytes} bytes.");
            }

            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
        }
    }

    private static string ResolveFileName(
        EmailAttachmentUrlOption option,
        HttpResponseMessage response,
        Uri uri)
    {
        if (!string.IsNullOrWhiteSpace(option.FileName))
        {
            return option.FileName.Trim();
        }

        var dispositionName = response.Content.Headers.ContentDisposition?.FileNameStar
            ?? response.Content.Headers.ContentDisposition?.FileName;
        if (!string.IsNullOrWhiteSpace(dispositionName))
        {
            return dispositionName.Trim().Trim('"');
        }

        var urlName = Path.GetFileName(uri.LocalPath);
        return string.IsNullOrWhiteSpace(urlName)
            ? $"adjunto-{Guid.NewGuid():N}.bin"
            : urlName;
    }

    private static string Truncate(string? value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value.Length <= maxLength ? value : value[..maxLength];
    }

    private sealed record AttachmentData(string Name, string ContentBase64);
}
