namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAsientoDetalleRequest
{
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public string Dh { get; init; } = "D";
    public string? GlosaDetalle { get; init; }
    public string? CodigoCentroCosto { get; init; }
    public string? TipoDocumento { get; init; }
    public string? NumeroDocumento { get; init; }
    public string? Serie { get; init; }
    public decimal TipoCambioLinea { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string? ReferenciaLinea { get; init; }
}
