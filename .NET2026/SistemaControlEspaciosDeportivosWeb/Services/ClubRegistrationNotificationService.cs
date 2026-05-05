using System.Net.Mail;
using System.Text;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class ClubRegistrationNotificationService(
    ISportCenterStoredProcedureService spService,
    IEmailService emailService,
    ILogger<ClubRegistrationNotificationService> logger) : IClubRegistrationNotificationService
{
    private const string ParametroCorreo1 = "HOME_PORTAL_NOTIF_CORREO_1";
    private const string ParametroCorreo2 = "HOME_PORTAL_NOTIF_CORREO_2";
    private const string SenderEmail = "info@lazonadeportiva.com";
    private const string SenderName = "La Zona Deportiva";

    public async Task NotifyNewClubRegistrationAsync(AltaClubSolicitudFormViewModel request, string? requestCode = null)
    {
        var recipients = await GetRecipientsAsync();
        if (recipients.Count == 0)
        {
            logger.LogInformation(
                "No se envio notificacion de alta de club porque {Parametro1} y {Parametro2} estan vacios.",
                ParametroCorreo1,
                ParametroCorreo2);
            return;
        }

        var subject = $"Nuevo registro de club/negocio: {SanitizeValue(request.NombreClub)}";
        var html = BuildNotificationTemplate(request, requestCode);

        foreach (var recipient in recipients)
        {
            try
            {
                await emailService.SendEmailAsync(
                    recipient,
                    recipient,
                    subject,
                    html,
                    new EmailSendOptions
                    {
                        SenderEmail = SenderEmail,
                        SenderName = SenderName
                    });
            }
            catch (EmailDeliveryException ex)
            {
                logger.LogWarning(ex, "No se pudo enviar notificacion de alta de club a {Recipient}.", recipient);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error inesperado al enviar notificacion de alta de club a {Recipient}.", recipient);
            }
        }
    }

    public async Task NotifyClubApprovalAsync(AltaClubItemViewModel request, int diasPrueba)
    {
        var correoDestino = (request.Correo ?? string.Empty).Trim();
        if (!IsValidEmail(correoDestino))
        {
            logger.LogInformation(
                "No se envio correo de aprobacion para solicitud {SolicitudId} porque no tiene correo valido.",
                request.Id);
            return;
        }

        var diasPruebaNormalizado = diasPrueba <= 0 ? 30 : diasPrueba;
        var fechaInicio = DateOnly.FromDateTime(DateTime.Today);
        var fechaFin = fechaInicio.AddDays(diasPruebaNormalizado);
        var asunto = $"Tu solicitud ha sido aprobada - Bienvenido a La Zona Deportiva";
        var html = BuildApprovalTemplate(request, diasPruebaNormalizado, fechaInicio, fechaFin);

        try
        {
            await emailService.SendEmailAsync(
                correoDestino,
                request.NombreContacto,
                asunto,
                html,
                new EmailSendOptions
                {
                    SenderEmail = SenderEmail,
                    SenderName = SenderName
                });
        }
        catch (EmailDeliveryException ex)
        {
            logger.LogWarning(
                ex,
                "No se pudo enviar correo de aprobacion a {Correo} para solicitud {SolicitudId}.",
                correoDestino,
                request.Id);
        }
        catch (Exception ex)
        {
            logger.LogError(
                ex,
                "Error inesperado al enviar correo de aprobacion a {Correo} para solicitud {SolicitudId}.",
                correoDestino,
                request.Id);
        }
    }

    private async Task<List<string>> GetRecipientsAsync()
    {
        var value1 = await spService.ParametrosGlobalesObtenerValorAsync(ParametroCorreo1);
        var value2 = await spService.ParametrosGlobalesObtenerValorAsync(ParametroCorreo2);

        var recipients = ParseEmails(value1)
            .Concat(ParseEmails(value2))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return recipients;
    }

    private static IEnumerable<string> ParseEmails(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            yield break;
        }

        var tokens = value
            .Split([';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var token in tokens)
        {
            if (IsValidEmail(token))
            {
                yield return token;
            }
        }
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            _ = new MailAddress(email);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string BuildNotificationTemplate(AltaClubSolicitudFormViewModel request, string? requestCode)
    {
        var fechaRegistro = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
        var codigoSolicitud = SanitizeValue(requestCode);
        var rows = new List<(string Label, string Value)>
        {
            ("Nombre del club", request.NombreClub),
            ("Nombre del usuario de registro", request.NombreContacto),
            ("Correo", request.Correo),
            ("Telefono", request.Telefono),
            ("Distrito", request.Ciudad),
            ("Provincia/Estado", request.ProvinciaEstado),
            ("Direccion", request.Direccion),
            ("Relacion con el club", request.RelacionClub),
            ("Fecha de registro", fechaRegistro)
        };

        if (!string.IsNullOrWhiteSpace(codigoSolicitud))
        {
            rows.Insert(0, ("Codigo de solicitud", codigoSolicitud));
        }

        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            sb.AppendLine(
                $"""
                <tr>
                  <td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">{Escape(row.Label)}</td>
                  <td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(SanitizeValue(row.Value))}</td>
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
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">Nuevo registro de club/negocio</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">
                  Se registro un nuevo club desde el portal publico. Revisa y da seguimiento a esta solicitud.
                </p>
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:16px;border:1px solid #dbe6f4;border-radius:10px;overflow:hidden;">
                  {sb}
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

    private static string BuildNotificationTemplate(AltaClubItemViewModel request)
    {
        var fechaRegistro = request.FechaRegistro.ToString("dd/MM/yyyy HH:mm");
        var rows = new List<(string Label, string Value)>
        {
            ("Codigo de solicitud", request.CodigoSolicitud),
            ("Nombre del club", request.NombreClub),
            ("Nombre del usuario de registro", request.NombreContacto),
            ("Correo", request.Correo),
            ("Telefono", request.Telefono),
            ("Distrito", request.Ciudad),
            ("Provincia/Estado", request.ProvinciaEstado),
            ("Direccion", request.Direccion),
            ("Relacion con el club", request.RelacionClub),
            ("Fecha de registro", fechaRegistro)
        };

        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            sb.AppendLine(
                $"""
                <tr>
                  <td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">{Escape(row.Label)}</td>
                  <td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(SanitizeValue(row.Value))}</td>
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
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">Nuevo registro de club/negocio</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">
                  Se registro un nuevo club desde el portal publico. Revisa y da seguimiento a esta solicitud.
                </p>
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:16px;border:1px solid #dbe6f4;border-radius:10px;overflow:hidden;">
                  {sb}
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

    private static string BuildApprovalTemplate(
        AltaClubItemViewModel request,
        int diasPrueba,
        DateOnly fechaInicio,
        DateOnly fechaFin)
    {
        var rows = new List<(string Label, string Value)>
        {
            ("Codigo de solicitud", request.CodigoSolicitud),
            ("Complejo Deportivo", request.NombreClub),
            ("Responsable", request.NombreContacto),
            ("Correo", request.Correo),
            ("Telefono", request.Telefono),
            ("Periodo de prueba", $"{diasPrueba} dias"),
            ("Inicio de prueba", fechaInicio.ToString("dd/MM/yyyy")),
            ("Fin de prueba", fechaFin.ToString("dd/MM/yyyy"))
        };

        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            sb.AppendLine(
                $"""
                <tr>
                  <td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">{Escape(row.Label)}</td>
                  <td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(SanitizeValue(row.Value))}</td>
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
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">Tu solicitud fue aprobada</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">
                  Bienvenido a La Zona Deportiva. Tu club ya fue aprobado y tiene periodo de prueba activo.
                </p>
                <p style="margin:0 0 8px;font-size:14px;line-height:1.6;color:#475569;">
                  Ya puedes iniciar sesion con tu cuenta y comenzar la configuracion de tu negocio.
                </p>
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:16px;border:1px solid #dbe6f4;border-radius:10px;overflow:hidden;">
                  {sb}
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
