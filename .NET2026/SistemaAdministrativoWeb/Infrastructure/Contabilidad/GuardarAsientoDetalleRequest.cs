namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAsientoDetalleRequest
{
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public string? GlosaDetalle { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string? ReferenciaLinea { get; init; }
}
