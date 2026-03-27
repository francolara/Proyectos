namespace SistemaControlEspaciosDeportivosWeb.Models;

public class Sede
{
    public int Id { get; set; }
    public int NegocioId { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Direccion { get; set; } = string.Empty;
    public string? Telefono { get; set; }
    public bool Activo { get; set; } = true;
    public DateTime FechaCreacion { get; set; } = DateTime.UtcNow;
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioCreacion { get; set; }
    public string? UsuarioActualizacion { get; set; }

    public Negocio? Negocio { get; set; }
    public ICollection<EspacioDeportivo> EspaciosDeportivos { get; set; } = new List<EspacioDeportivo>();
}
