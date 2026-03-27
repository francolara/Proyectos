using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Options;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class NotificacionEmailService(IOptions<EmailSettings> options) : INotificacionEmailService
{
    private readonly EmailSettings _settings = options.Value;

    public async Task<bool> EnviarSolicitudRecibidaAsync(SolicitudNotificacionEmailViewModel solicitud)
    {
        if (!_settings.Enabled) return false;
        if (string.IsNullOrWhiteSpace(solicitud.Correo)) return false;
        if (solicitud.NotificadoCliente) return false;
        if (string.IsNullOrWhiteSpace(_settings.Host) || string.IsNullOrWhiteSpace(_settings.From)) return false;

        using var message = new MailMessage
        {
            From = new MailAddress(_settings.From),
            Subject = $"Solicitud recibida - {solicitud.CodigoSolicitud}",
            Body = ConstruirMensaje(solicitud),
            IsBodyHtml = false
        };

        message.To.Add(solicitud.Correo);

        using var smtp = new SmtpClient(_settings.Host, _settings.Port)
        {
            EnableSsl = _settings.EnableSsl
        };

        if (!string.IsNullOrWhiteSpace(_settings.User))
        {
            smtp.Credentials = new NetworkCredential(_settings.User, _settings.Password);
        }

        await smtp.SendMailAsync(message);
        return true;
    }

    public async Task<bool> EnviarRecordatorioReservaAsync(ReservaRecordatorioPendienteViewModel reserva)
    {
        if (!_settings.Enabled) return false;
        if (string.IsNullOrWhiteSpace(reserva.Correo)) return false;
        if (string.IsNullOrWhiteSpace(_settings.Host) || string.IsNullOrWhiteSpace(_settings.From)) return false;

        using var message = new MailMessage
        {
            From = new MailAddress(_settings.From),
            Subject = $"Recordatorio de reserva - #{reserva.ReservaId}",
            Body = ConstruirMensajeRecordatorio(reserva),
            IsBodyHtml = false
        };

        message.To.Add(reserva.Correo);
        if (!string.IsNullOrWhiteSpace(reserva.CorreoNotificacion))
        {
            message.Bcc.Add(reserva.CorreoNotificacion);
        }

        using var smtp = new SmtpClient(_settings.Host, _settings.Port)
        {
            EnableSsl = _settings.EnableSsl
        };

        if (!string.IsNullOrWhiteSpace(_settings.User))
        {
            smtp.Credentials = new NetworkCredential(_settings.User, _settings.Password);
        }

        await smtp.SendMailAsync(message);
        return true;
    }

    private static string ConstruirMensaje(SolicitudNotificacionEmailViewModel solicitud)
    {
        return
$@"Hola {solicitud.NombreSolicitante},

Recibimos tu solicitud de reserva.

Codigo: {solicitud.CodigoSolicitud}
Sede: {solicitud.Sede}
Espacio: {solicitud.Espacio}
Fecha: {solicitud.Fecha:dd/MM/yyyy}
Horario: {solicitud.HoraInicio:HH\:mm} - {solicitud.HoraFin:HH\:mm}

Nuestro equipo te contactara para confirmar disponibilidad y pago.

Gracias.";
    }

    private static string ConstruirMensajeRecordatorio(ReservaRecordatorioPendienteViewModel reserva)
    {
        var contactoWhatsapp = string.IsNullOrWhiteSpace(reserva.WhatsappContacto)
            ? string.Empty
            : $@"
WhatsApp de contacto: {reserva.WhatsappContacto}";

        return
$@"Hola {reserva.Cliente},

Te recordamos tu reserva programada.

Reserva: #{reserva.ReservaId}
Sede: {reserva.Sede}
Espacio: {reserva.Espacio}
Fecha: {reserva.Fecha:dd/MM/yyyy}
Horario: {reserva.HoraInicio:HH\:mm} - {reserva.HoraFin:HH\:mm}
{contactoWhatsapp}

Gracias por usar SportCenter.";
    }
}
