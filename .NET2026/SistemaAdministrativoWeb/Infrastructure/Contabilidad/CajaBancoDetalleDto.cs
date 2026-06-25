namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CajaBancoDetalleDto
{
    public int? IdMovimientoBancoDetalle { get; init; }
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public int? IdPersona { get; init; }
    public string? ModuloOperacionComprobante { get; init; }
    public int? IdRegistroComprobante { get; init; }
    public decimal? ImporteAplicado { get; init; }
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public string NombrePersona { get; init; } = string.Empty;
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public bool RequiereCentroCosto { get; init; }
    public string? GlosaDetalle { get; init; }
    public string? CodigoCentroCosto { get; init; }
    public string? NumeroDocumento { get; init; }
    public string? TipoDocumento { get; init; }
    public string? Serie { get; init; }
    public string? ReferenciaLinea { get; init; }
    public decimal? TipoCambioLinea { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
}
