using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class UsuariosNegocioIndexViewModel : ModuloBaseViewModel
{
    public List<UsuarioNegocioItemViewModel> Usuarios { get; set; } = new();
    public UsuarioNegocioAsignarFormViewModel AsignarForm { get; set; } = new();
}

public class UsuarioNegocioItemViewModel
{
    public int UsuarioNegocioId { get; set; }
    public string UsuarioId { get; set; } = string.Empty;
    public string Nombres { get; set; } = string.Empty;
    public string Apellidos { get; set; } = string.Empty;
    public string Correo { get; set; } = string.Empty;
    public int RolNegocio { get; set; }
    public bool Activo { get; set; }
}

public class UsuarioNegocioAsignarFormViewModel
{
    public int NegocioId { get; set; }

    [Required(ErrorMessage = "Este campo es obligatorio.")]
    [EmailAddress(ErrorMessage = "Ingresa un correo electronico valido.")]
    public string Correo { get; set; } = string.Empty;

    [Range(1, 5, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int RolNegocio { get; set; } = 2;
}

public class UsuarioNegocioRolFormViewModel
{
    public int NegocioId { get; set; }
    public int UsuarioNegocioId { get; set; }
    [Range(1, 5, ErrorMessage = "El campo {0} debe estar entre {1} y {2}.")]
    public int RolNegocio { get; set; }
}

public class UsuarioNegocioPermisosViewModel : ModuloBaseViewModel
{
    public int UsuarioNegocioId { get; set; }
    public string UsuarioNombre { get; set; } = string.Empty;
    public string UsuarioCorreo { get; set; } = string.Empty;
    public List<UsuarioNegocioPermisoModuloViewModel> Modulos { get; set; } = new();
    public List<SelectListItem> Roles { get; set; } = new();
    public int RolUsuarioNegocio { get; set; }
}

public class UsuarioNegocioPermisoModuloViewModel
{
    public int ModuloSistemaId { get; set; }
    public string ModuloCodigo { get; set; } = string.Empty;
    public string ModuloNombre { get; set; } = string.Empty;
    public bool PuedeVer { get; set; }
    public bool PuedeCrear { get; set; }
    public bool PuedeEditar { get; set; }
    public bool PuedeEliminar { get; set; }
}
