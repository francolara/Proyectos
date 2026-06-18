namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoPreviewRequestViewModel
{
    public int? IdConfiguracionContabilizacion { get; set; }
    public DateOnly FechaContabilizacion { get; set; }
    public string? TipoComprobante { get; set; }
    public string? Serie { get; set; }
    public string? Numero { get; set; }
    public decimal BaseImponible { get; set; }
    public decimal Igv { get; set; }
    public decimal Isc { get; set; }
    public decimal OtrosTributos { get; set; }
    public decimal Redondeo { get; set; }
    public decimal ImporteTotal { get; set; }
    public List<AsientoPreviewDetalleRequestViewModel> Detalles { get; set; } = [];
}
