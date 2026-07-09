namespace SistemaAdministrativoWeb.ViewModels.Panel;

public sealed class PanelIndexViewModel
{
    public int IdEmpresa { get; init; }
    public string NombreEmpresa { get; init; } = string.Empty;
    public string? Ruc { get; init; }
    public string EstadoPanel { get; init; } = string.Empty;
    public string ResumenVigencia { get; init; } = string.Empty;
    public string PeriodoActual { get; init; } = string.Empty;
    public bool PeriodoCerrado { get; init; }
    public string EstadoPeriodo { get; init; } = string.Empty;
    public int TotalComprasPeriodo { get; init; }
    public decimal ImporteComprasPeriodo { get; init; }
    public int TotalVentasPeriodo { get; init; }
    public decimal ImporteVentasPeriodo { get; init; }
    public int TotalAsientosPeriodo { get; init; }
    public decimal ImporteAsientosPeriodo { get; init; }
    public int TotalCuentasCorrientesActivas { get; init; }
    public IReadOnlyCollection<PanelIndicadorViewModel> Indicadores { get; init; } = [];
    public IReadOnlyCollection<PanelSeriePeriodoViewModel> ComprasPorPeriodo { get; init; } = [];
    public IReadOnlyCollection<PanelSeriePeriodoViewModel> VentasPorPeriodo { get; init; } = [];
    public IReadOnlyCollection<PanelDistribucionItemViewModel> DistribucionRegistros { get; init; } = [];
    public IReadOnlyCollection<PanelModuloViewModel> Modulos { get; init; } = [];
}
