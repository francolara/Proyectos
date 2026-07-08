namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class RegistroVentasItemDto
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
    public decimal BaseImponible { get; init; }
    public decimal Descuento { get; init; }
    public decimal TotalExonerado { get; init; }
    public decimal TotalInafecto { get; init; }
    public decimal Igv { get; init; }
    public decimal Isc { get; init; }
    public decimal OtrosTributos { get; init; }
    public decimal Icbper { get; init; }
    public decimal Redondeo { get; init; }
    public decimal ImporteTotal { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
}
