namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class ContextoSuscripcionEmpresaDto
{
    public int IdEmpresa { get; init; }
    public int IdCuentaAdministradora { get; init; }
    public string CodigoEmpresa { get; init; } = string.Empty;
    public string RazonSocial { get; init; } = string.Empty;
    public string? NombreComercial { get; init; }
    public string Ruc { get; init; } = string.Empty;
    public bool EstadoEmpresa { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CorreoPrincipal { get; init; } = string.Empty;
    public string? TelefonoPrincipal { get; init; }
    public bool EstadoCuenta { get; init; }
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
}
