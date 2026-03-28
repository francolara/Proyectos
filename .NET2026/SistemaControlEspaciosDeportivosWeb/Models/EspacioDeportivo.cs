namespace SistemaControlEspaciosDeportivosWeb.Models;

public class EspacioDeportivo
{
    public int Id { get; set; }
    public int SedeId { get; set; }
    public int TipoDeporteId { get; set; }
    public int TipoSueloId { get; set; }
    public string Codigo { get; set; } = string.Empty;
    public string Nombre { get; set; } = string.Empty;
    public int Capacidad { get; set; }
    public bool TieneIluminacion { get; set; }
    public bool Techada { get; set; }
    public EstadoEspacioDeportivo Estado { get; set; } = EstadoEspacioDeportivo.Activo;
    public DateTime FechaCreacion { get; set; } = DateTime.UtcNow;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioCreacion { get; set; }
    public string? UsuarioActualizacion { get; set; }

    public Sede? Sede { get; set; }
    public TipoDeporte? TipoDeporte { get; set; }
    public TipoSuelo? TipoSuelo { get; set; }
    public ICollection<Tarifa> Tarifas { get; set; } = new List<Tarifa>();
    public ICollection<Reserva> Reservas { get; set; } = new List<Reserva>();
}
