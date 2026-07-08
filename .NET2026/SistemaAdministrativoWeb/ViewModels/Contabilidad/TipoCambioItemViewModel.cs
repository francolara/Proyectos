namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TipoCambioItemViewModel
{
    public int IdTipoCambio { get; init; }
    public DateOnly Fecha { get; init; }
    public string IdMoneda { get; init; } = string.Empty;
    public decimal Compra { get; init; }
    public decimal Venta { get; init; }
    public decimal CompraSbs { get; init; }
    public decimal VentaSbs { get; init; }
    public string Fuente { get; init; } = string.Empty;
    public bool Estado { get; init; }
}
