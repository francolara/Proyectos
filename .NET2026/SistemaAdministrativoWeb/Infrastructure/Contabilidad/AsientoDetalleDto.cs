namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoDetalleDto
{
    public int IdAsientoDetalle { get; init; }
    public int IdAsiento { get; init; }
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string? GlosaDetalle { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string? ReferenciaLinea { get; init; }
}
