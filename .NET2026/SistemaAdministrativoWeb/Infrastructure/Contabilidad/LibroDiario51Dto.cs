namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiario51Dto
{
    public string PeriodoPle { get; init; } = string.Empty;
    public string Cuo { get; init; } = string.Empty;
    public string CorrelativoAsiento { get; init; } = string.Empty;
    public string CorrelativoMovimiento { get; init; } = string.Empty;
    public string CodigoCuentaContable { get; init; } = string.Empty;
    public string CodigoUnidadOperacion { get; init; } = string.Empty;
    public string CodigoCentroCosto { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public string CodigoLibroRelacionado { get; init; } = string.Empty;
    public string TipoDocumentoEmisor { get; init; } = string.Empty;
    public string NumeroDocumentoEmisor { get; init; } = string.Empty;
    public string TipoComprobante { get; init; } = string.Empty;
    public string SerieComprobante { get; init; } = string.Empty;
    public string NumeroComprobante { get; init; } = string.Empty;
    public DateOnly FechaContable { get; init; }
    public DateOnly? FechaVencimiento { get; init; }
    public DateOnly FechaOperacion { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string GlosaReferencial { get; init; } = string.Empty;
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string InformacionComplementaria { get; init; } = string.Empty;
    public string EstadoOperacion { get; init; } = PleEstadoRegistroCatalogo.Vigente;
    public int NumeroAsiento { get; init; }
}
