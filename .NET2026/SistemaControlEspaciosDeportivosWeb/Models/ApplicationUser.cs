using Microsoft.AspNetCore.Identity;

namespace SistemaControlEspaciosDeportivosWeb.Models;

public class ApplicationUser : IdentityUser
{
    public string? Nombres { get; set; }
    public string? Apellidos { get; set; }
    public ICollection<UsuarioNegocio> UsuariosNegocio { get; set; } = new List<UsuarioNegocio>();
}

