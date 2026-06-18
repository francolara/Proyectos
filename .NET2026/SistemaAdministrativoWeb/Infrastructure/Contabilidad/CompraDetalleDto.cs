namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CompraDetalleDto
{
    public int IdCompraDetalle { get; init; }
    public int IdCompra { get; init; }
    public short Item { get; init; }
    public string Descripcion { get; init; } = string.Empty;
    public decimal Cantidad { get; init; }
    public decimal ValorUnitario { get; init; }
    public decimal ImporteBruto { get; init; }
}
