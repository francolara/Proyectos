namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoPreviewDetalleRequest
{
    public short Item { get; init; }
    public string? Descripcion { get; init; }
    public decimal Cantidad { get; init; }
    public decimal ValorUnitario { get; init; }
    public decimal ImporteBruto { get; init; }
}
