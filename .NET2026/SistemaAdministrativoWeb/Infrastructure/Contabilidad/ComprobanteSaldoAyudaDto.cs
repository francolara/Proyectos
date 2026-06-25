namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ComprobanteSaldoAyudaDto
{
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdRegistro { get; init; }
    public DateOnly FechaEmision { get; init; }
    public string NombrePersona { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public string TipoComprobante { get; init; } = string.Empty;
    public string DescripcionTipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public decimal Saldo { get; init; }
}
