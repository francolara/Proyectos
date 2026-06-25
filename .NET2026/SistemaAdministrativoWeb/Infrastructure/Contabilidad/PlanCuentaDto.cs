namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PlanCuentaDto
{
    public int IdPlanCuenta { get; init; }
    public int? IdPlanCuentaPadre { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public int NivelCuenta { get; init; }
    public string ColBalance { get; init; } = string.Empty;
    public string IdMoneda { get; init; } = string.Empty;
    public string TipoCambio { get; init; } = string.Empty;
    public bool AceptaMovimiento { get; init; }
    public bool EsUltimoNivel { get; init; }
    public bool RequiereCentroCosto { get; init; }
    public bool Estado { get; init; }
}
