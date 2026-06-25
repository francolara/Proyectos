namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionNotaCreditoPendienteItemViewModel
{
    public int IdRegistro { get; init; }
    public int IdMoneda { get; init; }
    public DateOnly FechaEmision { get; init; }
    public string TipoComprobante { get; init; } = string.Empty;
    public string DescripcionTipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal ImporteTotal { get; init; }
    public decimal Saldo { get; init; }
    public string EscenarioOperacion { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
    public string DocumentoTexto => $"{TipoComprobante} - {DescripcionTipoComprobante}";
    public string SerieNumeroTexto => $"{Serie}-{Numero}";
}
