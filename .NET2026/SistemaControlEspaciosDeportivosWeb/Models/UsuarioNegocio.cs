namespace SistemaControlEspaciosDeportivosWeb.Models;

public class UsuarioNegocio
{
    public int Id { get; set; }
    public string UsuarioId { get; set; } = string.Empty;
    public int NegocioId { get; set; }
    public RolNegocio RolNegocio { get; set; } = RolNegocio.Trabajador;
    public bool Activo { get; set; } = true;

    public ApplicationUser? Usuario { get; set; }
    public Negocio? Negocio { get; set; }
    public ICollection<UsuarioNegocioPermiso> Permisos { get; set; } = new List<UsuarioNegocioPermiso>();
}
