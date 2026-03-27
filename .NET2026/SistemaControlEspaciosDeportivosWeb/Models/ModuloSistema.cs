using System.ComponentModel.DataAnnotations;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class ModuloSistema
{
    public int Id { get; set; }

    [MaxLength(50)]
    public string Codigo { get; set; } = string.Empty;

    [MaxLength(120)]
    public string Nombre { get; set; } = string.Empty;

    public bool Activo { get; set; } = true;

    public ICollection<RolNegocioPermiso> RolesPermiso { get; set; } = new List<RolNegocioPermiso>();
    public ICollection<UsuarioNegocioPermiso> UsuariosPermiso { get; set; } = new List<UsuarioNegocioPermiso>();
}

