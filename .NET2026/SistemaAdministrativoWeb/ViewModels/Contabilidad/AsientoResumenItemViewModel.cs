namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoResumenItemViewModel
{
    public int IdAsiento { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string Periodo { get; init; } = string.Empty;
    public int NumeroAsiento { get; init; }
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaAsiento { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public decimal TotalImporteS { get; init; }
    public decimal TotalImporteD { get; init; }
    public string Estado { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
}
