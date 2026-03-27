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
