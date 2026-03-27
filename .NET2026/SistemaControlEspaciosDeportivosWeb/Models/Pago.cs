namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Pago
{
    public int Id { get; set; }
    public int ReservaId { get; set; }
    public DateTime FechaPago { get; set; } = DateTime.UtcNow;
    public decimal Monto { get; set; }
    public FormaPago FormaPago { get; set; }
    public string? NumeroOperacion { get; set; }
    public string? Observacion { get; set; }
    public DateTime FechaCreacion { get; set; } = DateTime.UtcNow;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioCreacion { get; set; }
    public string? UsuarioActualizacion { get; set; }

    public Reserva? Reserva { get; set; }
}
