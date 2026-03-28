namespace SistemaControlEspaciosDeportivosWeb.Models;

public class TipoSuelo
{
    public int Id { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public bool Activo { get; set; } = true;

    public ICollection<EspacioDeportivo> EspaciosDeportivos { get; set; } = new List<EspacioDeportivo>();
}

