namespace SistemaControlEspaciosDeportivosWeb.Models;

public class NivelDesafio
{
    public int IdNivel { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public bool Activo { get; set; }
    public int Orden { get; set; }
}
