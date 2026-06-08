using System.Text;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public static class ReservationCalendarIcsBuilder
{
    private const string ProductId = "-//La Zona Deportiva//Reservas//ES";

    public static byte[] Build(UsuarioPublicoReservaCalendarioViewModel reserva, DateTime utcNow)
    {
        var zonaPeru = ObtenerZonaHorariaPeru();
        var inicioLocal = reserva.Fecha.ToDateTime(reserva.HoraInicio, DateTimeKind.Unspecified);
        var finLocal = reserva.Fecha.ToDateTime(reserva.HoraFin, DateTimeKind.Unspecified);
        var inicioUtc = TimeZoneInfo.ConvertTimeToUtc(inicioLocal, zonaPeru);
        var finUtc = TimeZoneInfo.ConvertTimeToUtc(finLocal, zonaPeru);
        var descripcion = string.Join(
            "\n",
            "Reserva realizada en La Zona Deportiva.",
            $"Complejo: {reserva.NegocioNombre}",
            $"Sede: {reserva.SedeNombre}",
            $"Espacio: {reserva.EspacioNombre}",
            $"Codigo de reserva: {reserva.CodigoReserva}",
            $"Estado: {reserva.EstadoTexto}");

        var lines = new[]
        {
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            $"PRODID:{ProductId}",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            "BEGIN:VEVENT",
            $"UID:reserva-{reserva.ReservaId}@lazonadeportiva.com",
            $"DTSTAMP:{FormatUtc(utcNow)}",
            $"DTSTART:{FormatUtc(inicioUtc)}",
            $"DTEND:{FormatUtc(finUtc)}",
            $"SUMMARY:{EscapeText($"Reserva - {reserva.EspacioNombre}")}",
            $"LOCATION:{EscapeText(reserva.SedeDireccion ?? reserva.SedeNombre)}",
            $"DESCRIPTION:{EscapeText(descripcion)}",
            "STATUS:CONFIRMED",
            "END:VEVENT",
            "END:VCALENDAR"
        };

        return Encoding.UTF8.GetBytes(string.Join("\r\n", lines) + "\r\n");
    }

    public static string EscapeText(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        return value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace(";", "\\;", StringComparison.Ordinal)
            .Replace(",", "\\,", StringComparison.Ordinal)
            .Replace("\r\n", "\\n", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal)
            .Replace("\r", "\\n", StringComparison.Ordinal);
    }

    public static TimeZoneInfo ObtenerZonaHorariaPeru()
    {
        foreach (var timeZoneId in new[] { "SA Pacific Standard Time", "America/Lima" })
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(timeZoneId);
            }
            catch (TimeZoneNotFoundException)
            {
            }
            catch (InvalidTimeZoneException)
            {
            }
        }

        return TimeZoneInfo.Utc;
    }

    private static string FormatUtc(DateTime value)
    {
        return value.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'");
    }
}
