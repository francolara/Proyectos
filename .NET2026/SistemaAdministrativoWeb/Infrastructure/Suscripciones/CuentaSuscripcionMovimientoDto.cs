namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class CuentaSuscripcionMovimientoDto
{
    public int IdCuentaAdministradoraSuscripcionMovimiento { get; init; }
    public string TipoMovimiento { get; init; } = string.Empty;
    public string? TipoPlanAnterior { get; init; }
    public string? TipoPlanNuevo { get; init; }
    public string? EstadoSuscripcionAnterior { get; init; }
    public string EstadoSuscripcionNuevo { get; init; } = string.Empty;
    public bool? EsPruebaAnterior { get; init; }
    public bool EsPruebaNuevo { get; init; }
    public string? TipoCobroAnterior { get; init; }
    public string? TipoCobroNuevo { get; init; }
    public DateOnly? FechaInicioReferencia { get; init; }
    public DateOnly? FechaFinReferencia { get; init; }
    public int DiasGracia { get; init; }
    public int DiasExtra { get; init; }
    public int? EmpresasPermitidasAnterior { get; init; }
    public int? EmpresasPermitidasNuevo { get; init; }
    public int? UsuariosPermitidosAnterior { get; init; }
    public int? UsuariosPermitidosNuevo { get; init; }
    public string? Observacion { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string? UsuarioRegistro { get; init; }
}
