namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class CuentaSuscripcionResumenDto
{
    public int IdCuentaAdministradora { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CorreoPrincipal { get; init; } = string.Empty;
    public string? TelefonoPrincipal { get; init; }
    public bool EstadoCuenta { get; init; }
    public int? IdEmpresaPrincipal { get; init; }
    public string? CodigoEmpresaPrincipal { get; init; }
    public string? RazonSocialEmpresaPrincipal { get; init; }
    public string? NombreComercialEmpresaPrincipal { get; init; }
    public string? RucEmpresaPrincipal { get; init; }
    public int CantidadEmpresas { get; init; }
    public int? IdCuentaAdministradoraSuscripcion { get; init; }
    public string? TipoPlan { get; init; }
    public string? EstadoSuscripcion { get; init; }
    public bool EsPrueba { get; init; }
    public DateOnly? FechaInicioPrueba { get; init; }
    public DateOnly? FechaFinPrueba { get; init; }
    public DateOnly? FechaInicioPlan { get; init; }
    public DateOnly? FechaFinPlan { get; init; }
    public int? EmpresasPermitidas { get; init; }
    public int? UsuariosPermitidos { get; init; }
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public string? AspNetUserId { get; init; }
    public string? NombreCompleto { get; init; }
    public string? Telefono { get; init; }
    public string? Email { get; init; }
}
