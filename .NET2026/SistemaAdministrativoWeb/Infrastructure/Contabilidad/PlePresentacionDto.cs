namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PlePresentacionContextoDto
{
    public int? IdGeneracionPeriodo { get; init; }
    public bool Presentado { get; init; }
    public DateTime? FechaPresentacion { get; init; }
    public string UsuarioPresentacion { get; init; } = string.Empty;
    public bool ExistePresentacionAnterior { get; init; }
    public bool MesAnteriorPresentado { get; init; }
    public bool ExistePresentacionPosterior { get; init; }
    public string SnapshotUltimaPresentacion { get; init; } = string.Empty;
}

public sealed class PlePlanCuentaSnapshotItemDto
{
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string PeriodoInformado { get; init; } = string.Empty;
}

public sealed class PlePlanCuentaExportItemDto
{
    public string PeriodoPle { get; init; } = string.Empty;
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string EstadoOperacion { get; init; } = "1";
}

public sealed class PlePresentacionUpdateRequest
{
    public int IdEmpresa { get; init; }
    public int IdLibroElectronicoGeneracion { get; init; }
    public bool Presentado { get; init; }
    public string UsuarioPresentacion { get; init; } = string.Empty;
}
