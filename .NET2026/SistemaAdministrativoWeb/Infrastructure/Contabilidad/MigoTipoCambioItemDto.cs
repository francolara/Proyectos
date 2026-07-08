namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoTipoCambioItemDto
{
    public DateOnly Fecha { get; init; }
    public string Moneda { get; init; } = string.Empty;
    public decimal PrecioCompra { get; init; }
    public decimal PrecioVenta { get; init; }
}
