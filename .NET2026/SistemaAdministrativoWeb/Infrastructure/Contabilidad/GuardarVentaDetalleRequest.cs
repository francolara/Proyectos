namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarVentaDetalleRequest
{
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public int IdTipoAfectacionIGV { get; init; }
    public string Descripcion { get; init; } = string.Empty;
    public decimal Cantidad { get; init; }
    public decimal ValorUnitario { get; init; }
    public decimal ImporteBruto { get; init; }
}
