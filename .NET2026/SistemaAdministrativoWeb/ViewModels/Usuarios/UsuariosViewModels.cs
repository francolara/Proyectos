using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Usuarios;

public sealed class UsuariosIndexViewModel
{
    public int IdCuentaAdministradora { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public UsuarioCuentaFormViewModel Formulario { get; set; } = new();
    public IReadOnlyCollection<UsuarioCuentaListaItemViewModel> Usuarios { get; set; } = [];
    public IReadOnlyCollection<UsuarioEmpresaAsignacionViewModel> EmpresasDisponibles { get; set; } = [];
}

public sealed class UsuarioCuentaFormViewModel
{
    [Display(Name = "Nombre completo")]
    [Required(ErrorMessage = "Ingrese el nombre del usuario.")]
    [StringLength(180, ErrorMessage = "El nombre del usuario no puede exceder 180 caracteres.")]
    public string NombreCompleto { get; set; } = string.Empty;

    [Display(Name = "Correo")]
    [Required(ErrorMessage = "Ingrese el correo del usuario.")]
    [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
    [StringLength(256, ErrorMessage = "El correo no puede exceder 256 caracteres.")]
    public string Correo { get; set; } = string.Empty;

    [Display(Name = "Telefono")]
    [StringLength(30, ErrorMessage = "El telefono no puede exceder 30 caracteres.")]
    public string? Telefono { get; set; }

    [Display(Name = "Contrasena temporal")]
    [StringLength(100, MinimumLength = 6, ErrorMessage = "La contrasena temporal debe tener entre 6 y 100 caracteres.")]
    [DataType(DataType.Password)]
    public string? PasswordTemporal { get; set; }

    [Display(Name = "Rol de cuenta")]
    [Required(ErrorMessage = "Seleccione el rol del usuario.")]
    public string RolCuenta { get; set; } = "OPERADOR";

    [Display(Name = "Cuenta predeterminada")]
    public bool EsCuentaPredeterminada { get; set; }

    [Display(Name = "Empresa predeterminada")]
    public int? IdEmpresaPredeterminada { get; set; }

    public List<int> EmpresasSeleccionadas { get; set; } = [];
}

public sealed class UsuarioCuentaListaItemViewModel
{
    public int IdUsuarioCuentaAdministradora { get; set; }
    public string AspNetUserId { get; set; } = string.Empty;
    public string CorreoUsuario { get; set; } = string.Empty;
    public string NombreCompleto { get; set; } = string.Empty;
    public string? Telefono { get; set; }
    public string RolCuenta { get; set; } = string.Empty;
    public bool EsCuentaPredeterminada { get; set; }
    public int CantidadEmpresasAsignadas { get; set; }
    public string EmpresasAsignadas { get; set; } = string.Empty;
}

public sealed class UsuarioEmpresaAsignacionViewModel
{
    public int IdEmpresa { get; set; }
    public string CodigoEmpresa { get; set; } = string.Empty;
    public string RazonSocial { get; set; } = string.Empty;
    public bool Asignado { get; set; }
    public bool EsEmpresaPredeterminada { get; set; }
    public int? IdUsuarioEmpresa { get; set; }
}

public sealed class UsuarioPermisosViewModel
{
    public int IdCuentaAdministradora { get; set; }
    public int IdUsuarioCuentaAdministradora { get; set; }
    public string AspNetUserId { get; set; } = string.Empty;
    public string CorreoUsuario { get; set; } = string.Empty;
    public string NombreCompleto { get; set; } = string.Empty;
    public string RolCuenta { get; set; } = string.Empty;
    public bool EsCuentaPredeterminada { get; set; }
    public int? IdEmpresaPredeterminada { get; set; }
    public List<int> EmpresasSeleccionadas { get; set; } = [];
    public IReadOnlyCollection<UsuarioEmpresaAsignacionViewModel> EmpresasDisponibles { get; set; } = [];
    public IReadOnlyCollection<ModuloPermisoViewModel> PermisosCuenta { get; set; } = [];
    public IReadOnlyCollection<ModuloPermisoViewModel> PermisosEmpresa { get; set; } = [];
    public int? IdUsuarioEmpresaSeleccionada { get; set; }
    public int? IdEmpresaSeleccionada { get; set; }
    public string? NombreEmpresaSeleccionada { get; set; }
    public string PestanaActiva { get; set; } = "general";
}

public sealed class ModuloPermisoViewModel
{
    public int IdModuloSistema { get; set; }
    public string CodigoModulo { get; set; } = string.Empty;
    public string NombreModulo { get; set; } = string.Empty;
    public string GrupoMenu { get; set; } = string.Empty;
    public bool? PuedeVerOverride { get; set; }
    public bool? PuedeCrearOverride { get; set; }
    public bool? PuedeEditarOverride { get; set; }
    public bool? PuedeEliminarOverride { get; set; }
    public bool PuedeVerRol { get; set; }
    public bool PuedeCrearRol { get; set; }
    public bool PuedeEditarRol { get; set; }
    public bool PuedeEliminarRol { get; set; }
    public bool PuedeVerEfectivo { get; set; }
    public bool PuedeCrearEfectivo { get; set; }
    public bool PuedeEditarEfectivo { get; set; }
    public bool PuedeEliminarEfectivo { get; set; }
}

public sealed class GuardarPermisosLoteViewModel
{
    [Required]
    public string AspNetUserId { get; set; } = string.Empty;

    public int IdUsuarioCuentaAdministradora { get; set; }
    public int? IdUsuarioEmpresa { get; set; }
    public int? IdEmpresa { get; set; }
    public List<ModuloPermisoEdicionViewModel> Permisos { get; set; } = [];
}

public sealed class ModuloPermisoEdicionViewModel
{
    public int IdModuloSistema { get; set; }
    public bool? PuedeVer { get; set; }
    public bool? PuedeCrear { get; set; }
    public bool? PuedeEditar { get; set; }
    public bool? PuedeEliminar { get; set; }
}
