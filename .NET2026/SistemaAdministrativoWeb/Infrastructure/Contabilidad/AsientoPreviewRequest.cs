namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoPreviewRequest
{
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdConfiguracionContabilizacion { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public string TipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public decimal BaseImponible { get; init; }
    public decimal Igv { get; init; }
    public decimal Isc { get; init; }
    public decimal OtrosTributos { get; init; }
    public decimal Redondeo { get; init; }
    public decimal ImporteTotal { get; init; }
    public IReadOnlyCollection<AsientoPreviewDetalleRequest> Detalles { get; init; } = [];
}
