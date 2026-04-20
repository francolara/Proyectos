namespace SistemaControlEspaciosDeportivosWeb.Models;

public class DesafioMensaje
{
    public int IdMensaje { get; set; }
    public int IdDesafio { get; set; }
    public string UsuarioIdEmisor { get; set; } = string.Empty;
    public string Mensaje { get; set; } = string.Empty;
    public DateTime FechaRegistro { get; set; }
    public bool Activo { get; set; }
}
