using System.Text;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using Microsoft.AspNetCore.Http;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class ReservationEmailNotificationService(
    ISportCenterStoredProcedureService spService,
    IEmailService emailService,
    ILogger<ReservationEmailNotificationService> logger,
    IHostEnvironment environment,
    IHttpContextAccessor httpContextAccessor) : IReservationEmailNotificationService
{
    private const string SenderEmailReservas = "reservas@lazonadeportiva.com";
    private const string SenderNameReservas = "La Zona Deportiva";

    public async Task NotifyPublicReservationCreatedAsync(int? negocioId, int reservaId)
    {
        logger.LogInformation(
            "Inicio notificacion de reserva publica. NegocioId={NegocioId}, ReservaId={ReservaId}.",
            negocioId,
            reservaId);
        try
        {
            var reserva = await spService.ReservasObtenerContextoEmailAsync( negocioId, reservaId);
            if (reserva is null && negocioId.HasValue)
            {
                // Fallback defensivo: para reserva publica priorizamos encontrar por ReservaId
                // aunque el NegocioId posteado no coincida.
                reserva = await spService.ReservasObtenerContextoEmailAsync(null, reservaId);
            }
            if (reserva is null) return;

            if (!reserva.NotificacionesActivasSede)
            {
                logger.LogInformation(
                    "No se envio correo de reserva publica {ReservaId} porque la sede tiene notificaciones desactivadas.",
                    reservaId);
                return;
            }

            if (string.IsNullOrWhiteSpace(reserva.CorreoNotificacionSede))
            {
                logger.LogInformation(
                    "No se envio correo de reserva publica {ReservaId} porque la sede no tiene correo de notificacion.",
                    reservaId);
                return;
            }

            await emailService.SendEmailAsync(
                reserva.CorreoNotificacionSede.Trim(),
                reserva.Sede,
                $"Nueva reserva generada - #{reserva.ReservaId:D6}",
                ReservationEmailTemplateBuilder.BuildReservationGeneratedTemplate(
                    reserva,
                    BuildLoginReturnUrl("/Reservas"),
                    "Gestionar reserva"),
                new EmailSendOptions
                {
                    SenderEmail = SenderEmailReservas,
                    SenderName = SenderNameReservas
                });

            logger.LogInformation(
                "Correo de reserva publica enviado. ReservaId={ReservaId}, CorreoDestino={CorreoDestino}.",
                reservaId,
                reserva.CorreoNotificacionSede.Trim());
        }
        catch (Exception ex)
        {
            logger.LogWarning(
                ex,
                "No se pudo enviar notificacion de reserva publica. ReservaId={ReservaId}, NegocioId={NegocioId}.",
                reservaId,
                negocioId);

            if (environment.IsDevelopment())
            {
                throw;
            }
        }
    }

    public async Task NotifyReservationConfirmedIfAppliesAsync(int negocioId, int reservaId, int? estadoAnterior = null)
    {
        logger.LogInformation(
            "Inicio notificacion de confirmacion. NegocioId={NegocioId}, ReservaId={ReservaId}, EstadoAnterior={EstadoAnterior}.",
            negocioId,
            reservaId,
            estadoAnterior);
        try
        {
            var reserva = await spService.ReservasObtenerContextoEmailAsync(negocioId, reservaId);
            if (reserva is null) return;

            if (!EsEstadoConfirmadoOPagado(reserva.Estado))
            {
                return;
            }

            if (estadoAnterior.HasValue && EsEstadoConfirmadoOPagado(estadoAnterior.Value))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(reserva.ClienteCorreo))
            {
                logger.LogInformation(
                    "No se envio confirmacion de reserva {ReservaId} porque el cliente no tiene correo configurado.",
                    reservaId);
                return;
            }

            await emailService.SendEmailAsync(
                reserva.ClienteCorreo.Trim(),
                reserva.Cliente,
                $"Tu reserva fue confirmada - #{reserva.ReservaId:D6}",
                ReservationEmailTemplateBuilder.BuildReservationConfirmedTemplate(
                    reserva,
                    BuildLoginReturnUrl("/PerfilPublico?tab=reservas"),
                    "Ver reserva"),
                new EmailSendOptions
                {
                    SenderEmail = SenderEmailReservas,
                    SenderName = SenderNameReservas
                });

            logger.LogInformation(
                "Correo de confirmacion enviado. ReservaId={ReservaId}, CorreoDestino={CorreoDestino}.",
                reservaId,
                reserva.ClienteCorreo.Trim());
        }
        catch (Exception ex)
        {
            logger.LogWarning(
                ex,
                "No se pudo enviar correo de confirmacion. ReservaId={ReservaId}, NegocioId={NegocioId}.",
                reservaId,
                negocioId);

            if (environment.IsDevelopment())
            {
                throw;
            }
        }
    }

    private static bool EsEstadoConfirmadoOPagado(int estado)
    {
        return estado == (int)EstadoReserva.Confirmada || estado == (int)EstadoReserva.Pagada;
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

public static class ReservationEmailTemplateBuilder
{
    public static string BuildReservationGeneratedTemplate(
        ReservaEmailContextViewModel reserva,
        string actionUrl,
        string actionText)
    {
        return BuildBaseTemplate(
            title: "Reserva generada",
            intro: "Se registro una nueva reserva desde el Home.",
            body: "Revisa la informacion para gestionar el seguimiento de esta reserva.",
            reserva: reserva,
            actionUrl: actionUrl,
            actionText: actionText);
    }

    public static string BuildReservationConfirmedTemplate(
        ReservaEmailContextViewModel reserva,
        string actionUrl,
        string actionText)
    {
        var estadoTexto = reserva.Estado == (int)EstadoReserva.Pagada ? "pagada" : "confirmada";
        return BuildBaseTemplate(
            title: "Reserva confirmada",
            intro: $"Tu reserva fue {estadoTexto} correctamente.",
            body: "Te compartimos el detalle para tu control.",
            reserva: reserva,
            actionUrl: actionUrl,
            actionText: actionText);
    }

    private static string BuildBaseTemplate(
        string title,
        string intro,
        string body,
        ReservaEmailContextViewModel reserva,
        string actionUrl,
        string actionText)
    {
        var rows = new List<(string Label, string Value)>
        {
            ("Club / negocio", SanitizeValue(reserva.Negocio)),
            ("Codigo de reserva", $"#{reserva.ReservaId:D6}"),
            ("Sede", SanitizeValue(reserva.Sede)),
            ("Espacio deportivo", SanitizeValue(reserva.Espacio)),
            ("Fecha", reserva.Fecha.ToString("dd/MM/yyyy")),
            ("Horario", $"{reserva.HoraInicio:HH\\:mm} - {reserva.HoraFin:HH\\:mm}"),
            ("Cliente", SanitizeValue(reserva.Cliente)),
            ("Equipo", SanitizeValue(reserva.NombreEquipo)),
            ("Telefono", SanitizeValue(reserva.ClienteTelefono)),
            ("Correo", SanitizeValue(reserva.ClienteCorreo))
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
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">{Escape(title)}</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">{Escape(intro)}</p>
                <p style="margin:0 0 8px;font-size:14px;line-height:1.6;color:#475569;">{Escape(body)}</p>
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
