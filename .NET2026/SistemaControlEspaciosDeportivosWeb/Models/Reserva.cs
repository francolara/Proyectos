namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Reserva
{
    public int Id { get; set; }
    public int EspacioDeportivoId { get; set; }
    public int ClienteId { get; set; }
    public DateOnly Fecha { get; set; }
    public TimeOnly HoraInicio { get; set; }
    public TimeOnly HoraFin { get; set; }
    public EstadoReserva Estado { get; set; } = EstadoReserva.Pendiente;
    public decimal Total { get; set; }
    public decimal Adelanto { get; set; }
    public decimal Saldo { get; set; }
    public DateTime FechaRegistro { get; set; } = DateTime.UtcNow;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioCreacion { get; set; }
    public string? UsuarioActualizacion { get; set; }

    public EspacioDeportivo? EspacioDeportivo { get; set; }
    public Cliente? Cliente { get; set; }
    public ICollection<Pago> Pagos { get; set; } = new List<Pago>();
    public ComprobanteElectronico? ComprobanteElectronico { get; set; }
}
