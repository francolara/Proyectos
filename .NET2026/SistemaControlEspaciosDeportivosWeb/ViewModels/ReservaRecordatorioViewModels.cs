namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ReservaRecordatorioPendienteViewModel
{
    public int ReservaId { get; set; }
    public int NegocioId { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string Correo { get; set; } = string.Empty;
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public string? CorreoNotificacion { get; set; }
    public string? WhatsappContacto { get; set; }
}

public class ReservaEmailContextViewModel
{
    public int ReservaId { get; set; }
    public int NegocioId { get; set; }
    public string Negocio { get; set; } = string.Empty;
    public int Estado { get; set; }
    public string Cliente { get; set; } = string.Empty;
    public string? ClienteCorreo { get; set; }
    public string? ClienteTelefono { get; set; }
    public string? NombreEquipo { get; set; }
    public string Sede { get; set; } = string.Empty;
    public string Espacio { get; set; } = string.Empty;
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public bool NotificacionesActivasSede { get; set; }
    public string? CorreoNotificacionSede { get; set; }
}
