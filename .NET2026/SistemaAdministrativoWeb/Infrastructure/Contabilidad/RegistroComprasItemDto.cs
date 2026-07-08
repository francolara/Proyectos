namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class RegistroComprasItemDto
{
    public DateOnly FechaEmision { get; init; }
    public DateOnly FechaContabilizacion { get; init; }
    public string TipoComprobante { get; init; } = string.Empty;
    public string DescripcionTipoComprobante { get; init; } = string.Empty;
    public string Serie { get; init; } = string.Empty;
    public string Numero { get; init; } = string.Empty;
    public string CodigoPersona { get; init; } = string.Empty;
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public string NombrePersona { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal BaseImponibleGravada { get; init; }
    public decimal IgvGravado { get; init; }
    public decimal BaseImponibleGasto { get; init; }
    public decimal IgvGasto { get; init; }
    public decimal BaseImponibleSinCredito { get; init; }
    public decimal IgvSinCredito { get; init; }
    public decimal TotalExonerado { get; init; }
    public decimal TotalInafecto { get; init; }
    public decimal OtrosTributos { get; init; }
    public decimal Icbper { get; init; }
    public decimal Retencion { get; init; }
    public decimal ImporteDetraccion { get; init; }
    public decimal ImportePercepcion { get; init; }
    public decimal ImporteTotal { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
}
