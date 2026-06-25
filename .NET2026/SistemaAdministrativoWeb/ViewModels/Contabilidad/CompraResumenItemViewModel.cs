namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CompraResumenItemViewModel
{
    public int IdCompra { get; init; }
    public string NombreProveedor { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public string Documento { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public decimal Saldo { get; init; }
    public int? IdAsiento { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Situacion { get; init; } = string.Empty;
}
