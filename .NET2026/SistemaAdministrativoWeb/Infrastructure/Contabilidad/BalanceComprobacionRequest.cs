namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BalanceComprobacionRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public byte PeriodoDesde { get; init; }
    public byte PeriodoHasta { get; init; }
    public string Moneda { get; init; } = "PEN";
    public byte Grado { get; init; } = 1;
    public bool TodasLasCuentas { get; init; } = true;
    public string CuentaDesde { get; init; } = string.Empty;
    public string CuentaHasta { get; init; } = string.Empty;
    public bool FiltrarGrado { get; init; } = true;
}
