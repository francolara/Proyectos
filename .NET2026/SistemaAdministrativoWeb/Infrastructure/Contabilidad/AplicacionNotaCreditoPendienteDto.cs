namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AplicacionNotaCreditoPendienteDto
{
    public int IdRegistro { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdPersona { get; init; }
    public string NombrePersona { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public string TipoComprobante { get; init; } = string.Empty;
    public string DescripcionTipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal ImporteTotal { get; init; }
    public decimal Saldo { get; init; }
    public bool EsNotaCredito { get; init; }
    public string EscenarioOperacion { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
}
