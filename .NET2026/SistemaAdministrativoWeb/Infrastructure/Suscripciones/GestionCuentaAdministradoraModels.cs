namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class ContextoLoginUsuarioDto
{
    public bool TieneAcceso { get; set; }
    public bool EsSuperAdmin { get; set; }
    public int? IdCuentaAdministradora { get; set; }
    public string? CodigoCuenta { get; set; }
    public string? NombreCuenta { get; set; }
    public string? CorreoPrincipal { get; set; }
    public string? TelefonoPrincipal { get; set; }
    public bool? EstadoCuenta { get; set; }
    public string? RolCuenta { get; set; }
    public int CantidadEmpresasAsignadas { get; set; }
    public int? IdEmpresaPredeterminada { get; set; }
    public string? RazonSocialEmpresaPredeterminada { get; set; }
    public bool DebeSeleccionarEmpresa { get; set; }
    public bool SoloModulosCuenta { get; set; }
    public int? IdCuentaAdministradoraSuscripcion { get; set; }
    public string? TipoPlan { get; set; }
    public string? EstadoSuscripcion { get; set; }
    public bool? EsPrueba { get; set; }
    public DateOnly? FechaInicioPrueba { get; set; }
    public DateOnly? FechaFinPrueba { get; set; }
    public DateOnly? FechaInicioPlan { get; set; }
    public DateOnly? FechaFinPlan { get; set; }
    public int? DiasGracia { get; set; }
    public DateOnly? FechaFinGracia { get; set; }
    public int? EmpresasPermitidas { get; set; }
    public int? UsuariosPermitidos { get; set; }
    public bool? ActivoSuscripcion { get; set; }
    public string? ObservacionSuscripcion { get; set; }
    public string? Mensaje { get; set; }
}

public sealed class CuentaAdministradoraConfiguracionDto
{
    public int IdCuentaAdministradora { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string? CorreoPrincipal { get; set; }
    public string? TelefonoPrincipal { get; set; }
    public int? IdCuentaAdministradoraConfiguracion { get; set; }
    public string? NombreResponsablePrincipal { get; set; }
    public string? CorreoAdministrativo { get; set; }
    public string? TelefonoAdministrativo { get; set; }
    public int? IdEmpresaPredeterminada { get; set; }
    public string? CodigoEmpresaPredeterminada { get; set; }
    public string? RazonSocialEmpresaPredeterminada { get; set; }
    public string? ObservacionAdministrativa { get; set; }
    public int? IdCuentaAdministradoraFacturacion { get; set; }
    public string? TipoComprobantePreferido { get; set; }
    public string? TipoDocumentoFacturacion { get; set; }
    public string? NumeroDocumento { get; set; }
    public string? NombreFacturacion { get; set; }
    public string? RazonSocialFacturacion { get; set; }
    public string? CorreoFacturacion { get; set; }
    public string? TelefonoFacturacion { get; set; }
    public string? DireccionFiscal { get; set; }
    public string? Ubigeo { get; set; }
    public string? Distrito { get; set; }
    public string? Provincia { get; set; }
    public string? Departamento { get; set; }
    public string? ObservacionFacturacion { get; set; }
}

public sealed class CuentaAdministradoraEmpresaDto
{
    public int IdEmpresa { get; set; }
    public string CodigoEmpresa { get; set; } = string.Empty;
    public string RazonSocial { get; set; } = string.Empty;
    public string? NombreComercial { get; set; }
    public string? Ruc { get; set; }
    public bool Estado { get; set; }
}

public sealed class CuentaAdministradoraUsuarioDto
{
    public int IdUsuarioCuentaAdministradora { get; set; }
    public string AspNetUserId { get; set; } = string.Empty;
    public string CorreoUsuario { get; set; } = string.Empty;
    public string? NombreCompleto { get; set; }
    public string? Telefono { get; set; }
    public string RolCuenta { get; set; } = string.Empty;
    public bool EsCuentaPredeterminada { get; set; }
    public bool Estado { get; set; }
    public int CantidadEmpresasAsignadas { get; set; }
    public string? EmpresasAsignadas { get; set; }
}

public sealed class UsuarioCuentaEmpresaAsignadaDto
{
    public int IdEmpresa { get; set; }
    public string CodigoEmpresa { get; set; } = string.Empty;
    public string RazonSocial { get; set; } = string.Empty;
    public string? NombreComercial { get; set; }
    public string? Ruc { get; set; }
    public bool Asignado { get; set; }
    public bool EsEmpresaPredeterminada { get; set; }
    public int? IdUsuarioEmpresa { get; set; }
}

public sealed class UsuarioCuentaPermisoDto
{
    public int IdModuloSistema { get; set; }
    public string CodigoModulo { get; set; } = string.Empty;
    public string NombreModulo { get; set; } = string.Empty;
    public string? GrupoMenu { get; set; }
    public string? RolCuenta { get; set; }
    public bool? PuedeVerRol { get; set; }
    public bool? PuedeCrearRol { get; set; }
    public bool? PuedeEditarRol { get; set; }
    public bool? PuedeEliminarRol { get; set; }
    public bool? PuedeVerOverride { get; set; }
    public bool? PuedeCrearOverride { get; set; }
    public bool? PuedeEditarOverride { get; set; }
    public bool? PuedeEliminarOverride { get; set; }
    public bool PuedeVerEfectivo { get; set; }
    public bool PuedeCrearEfectivo { get; set; }
    public bool PuedeEditarEfectivo { get; set; }
    public bool PuedeEliminarEfectivo { get; set; }
}

public sealed class GuardarConfiguracionCuentaAdministradoraRequest
{
    public int IdCuentaAdministradora { get; set; }
    public string? NombreResponsablePrincipal { get; set; }
    public string? CorreoAdministrativo { get; set; }
    public string? TelefonoAdministrativo { get; set; }
    public int? IdEmpresaPredeterminada { get; set; }
    public string? ObservacionAdministrativa { get; set; }
    public string TipoComprobantePreferido { get; set; } = "BOLETA";
    public string TipoDocumentoFacturacion { get; set; } = "DNI";
    public string? NumeroDocumento { get; set; }
    public string? NombreFacturacion { get; set; }
    public string? RazonSocialFacturacion { get; set; }
    public string? CorreoFacturacion { get; set; }
    public string? TelefonoFacturacion { get; set; }
    public string? DireccionFiscal { get; set; }
    public string? Ubigeo { get; set; }
    public string? Distrito { get; set; }
    public string? Provincia { get; set; }
    public string? Departamento { get; set; }
    public string? ObservacionFacturacion { get; set; }
    public string? UsuarioRegistro { get; set; }
}

public sealed class AsignarUsuarioCuentaAdministradoraRequest
{
    public string AspNetUserId { get; set; } = string.Empty;
    public int IdCuentaAdministradora { get; set; }
    public string RolCuenta { get; set; } = "OPERADOR";
    public bool EsCuentaPredeterminada { get; set; }
    public string? UsuarioRegistro { get; set; }
}

public sealed class AsignarUsuarioEmpresaRequest
{
    public string AspNetUserId { get; set; } = string.Empty;
    public int IdEmpresa { get; set; }
    public bool EsEmpresaPredeterminada { get; set; }
    public string UsuarioRegistro { get; set; } = string.Empty;
}

public sealed class GuardarUsuarioCuentaPermisoRequest
{
    public int IdUsuarioCuentaAdministradora { get; set; }
    public int IdModuloSistema { get; set; }
    public bool? PuedeVer { get; set; }
    public bool? PuedeCrear { get; set; }
    public bool? PuedeEditar { get; set; }
    public bool? PuedeEliminar { get; set; }
    public string? UsuarioRegistro { get; set; }
}

public sealed class GuardarUsuarioEmpresaPermisoRequest
{
    public int IdUsuarioEmpresa { get; set; }
    public int IdModuloSistema { get; set; }
    public bool? PuedeVer { get; set; }
    public bool? PuedeCrear { get; set; }
    public bool? PuedeEditar { get; set; }
    public bool? PuedeEliminar { get; set; }
    public string? UsuarioRegistro { get; set; }
}
