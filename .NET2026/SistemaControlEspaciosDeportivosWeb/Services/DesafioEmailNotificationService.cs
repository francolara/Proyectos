using System.Text;
using Microsoft.AspNetCore.Http;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class DesafioEmailNotificationService(
    ISportCenterStoredProcedureService spService,
    IEmailService emailService,
    ILogger<DesafioEmailNotificationService> logger,
    IHostEnvironment environment,
    IHttpContextAccessor httpContextAccessor) : IDesafioEmailNotificationService
{
    private const string SenderEmailDesafios = "info@lazonadeportiva.com";
    private const string SenderNameDesafios = "La Zona Deportiva";

    public async Task NotifyDesafioReceivedAsync(int desafioId)
    {
        logger.LogInformation(
            "Inicio notificacion de desafio recibido. DesafioId={DesafioId}.",
            desafioId);
        try
        {
            var contexto = await spService.DesafiosObtenerContextoEmailAsync(desafioId);
            if (contexto is null)
            {
                logger.LogInformation(
                    "No se envio correo de desafio {DesafioId} porque no se encontro contexto.",
                    desafioId);
                return;
            }

            if (string.IsNullOrWhiteSpace(contexto.CorreoRetado))
            {
                logger.LogInformation(
                    "No se envio correo de desafio {DesafioId} porque el retado no tiene correo.",
                    desafioId);
                return;
            }

            await emailService.SendEmailAsync(
                contexto.CorreoRetado.Trim(),
                contexto.NombreRetado,
                $"Desafio recibido - #{contexto.DesafioId:D6}",
                DesafioEmailTemplateBuilder.BuildDesafioReceivedTemplate(
                    contexto,
                    BuildLoginReturnUrl("/Desafios"),
                    "Revisar desafio"),
                new EmailSendOptions
                {
                    SenderEmail = SenderEmailDesafios,
                    SenderName = SenderNameDesafios
                });

            logger.LogInformation(
                "Correo de desafio recibido enviado. DesafioId={DesafioId}, CorreoDestino={CorreoDestino}.",
                desafioId,
                contexto.CorreoRetado.Trim());
        }
        catch (Exception ex)
        {
            logger.LogWarning(
                ex,
                "No se pudo enviar notificacion de desafio recibido. DesafioId={DesafioId}.",
                desafioId);

            if (environment.IsDevelopment())
            {
                throw;
            }
        }
    }

    private string BuildLoginReturnUrl(string returnUrl)
    {
        var encodedReturnUrl = Uri.EscapeDataString(returnUrl);
        var loginPath = $"/Identity/Account/Login?returnUrl={encodedReturnUrl}";

        var request = httpContextAccessor.HttpContext?.Request;
        if (request is null || !request.Host.HasValue)
        {
            return loginPath;
        }

        var baseUrl = $"{request.Scheme}://{request.Host}{request.PathBase}";
        return $"{baseUrl.TrimEnd('/')}{loginPath}";
    }
}

public static class DesafioEmailTemplateBuilder
{
    public static string BuildDesafioReceivedTemplate(
        DesafioEmailContextViewModel desafio,
        string actionUrl,
        string actionText)
    {
        var rows = new List<(string Label, string Value)>
        {
            ("Codigo de desafio", $"#{desafio.DesafioId:D6}"),
            ("Equipo desafiante", SanitizeValue(desafio.EquipoRetador)),
            ("Contacto desafiante", SanitizeValue(desafio.ContactoRetador)),
            ("Usuario desafiante", SanitizeValue(desafio.UsuarioRetador)),
            ("Telefono desafiante", SanitizeValue(desafio.TelefonoRetador)),
            ("Deporte", SanitizeValue(desafio.Deporte)),
            ("Nivel", SanitizeValue(desafio.Nivel)),
            ("Distrito", SanitizeValue(desafio.Distrito)),
            ("Fecha tentativa", desafio.FechaTentativa.ToString("dd/MM/yyyy")),
            ("Hora tentativa", desafio.HoraTentativa.ToString(@"HH\:mm")),
            ("Cancha sugerida", SanitizeValue(desafio.CanchaSugerida)),
            ("Modalidad", SanitizeValue(desafio.Modalidad)),
            ("Forma de pago", SanitizeValue(desafio.FormaPago)),
            ("Mensaje", SanitizeValue(desafio.Mensaje))
        };

        var rowsHtml = new StringBuilder();
        foreach (var row in rows)
        {
            rowsHtml.AppendLine(
                $"""
                <tr>
                  <td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">{Escape(row.Label)}</td>
                  <td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(row.Value)}</td>
                </tr>
                """);
        }

        return
$"""
<!doctype html>
<html lang="es">
  <body style="margin:0;padding:0;background-color:#f4f7fb;font-family:Manrope,'Segoe UI',Arial,sans-serif;color:#1f2937;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color:#f4f7fb;padding:24px 12px;">
      <tr>
        <td align="center">
          <table role="presentation" width="680" cellspacing="0" cellpadding="0" style="max-width:680px;background:#ffffff;border:1px solid #dbe6f4;border-radius:14px;overflow:hidden;">
            <tr>
              <td style="background:linear-gradient(135deg,#0d3b66 0%,#164f86 60%,#17a2b8 100%);padding:22px 24px;">
                <div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#dbeafe;font-weight:800;">La Zona Deportiva</div>
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">Desafio recibido</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">
                  Recibiste un nuevo desafio. Te compartimos la informacion del equipo desafiante y los datos propuestos del partido.
                </p>
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:16px;border:1px solid #dbe6f4;border-radius:10px;overflow:hidden;">
                  {rowsHtml}
                </table>
                <table role="presentation" cellspacing="0" cellpadding="0" style="margin-top:18px;">
                  <tr>
                    <td style="border-radius:10px;background:linear-gradient(135deg,#0d3b66 0%,#164f86 60%,#17a2b8 100%);">
                      <a href="{Escape(actionUrl)}" style="display:inline-block;padding:12px 22px;color:#ffffff;text-decoration:none;font-size:14px;font-weight:700;border-radius:10px;">
                        {Escape(actionText)}
                      </a>
                    </td>
                  </tr>
                </table>
                <p style="margin:18px 0 0;font-size:12px;line-height:1.5;color:#94a3b8;">
                  Este correo fue enviado automaticamente por La Zona Deportiva.
                </p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
""";
    }

    private static string SanitizeValue(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
    }

    private static string Escape(string value)
    {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }
}
