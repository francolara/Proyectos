namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VentaResumenItemViewModel
{
    public int IdVenta { get; init; }
    public string NombreCliente { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public DateOnly FechaContabilizacion { get; init; }
    public string Documento { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public int? IdAsiento { get; init; }
    public string Estado { get; init; } = string.Empty;
}
