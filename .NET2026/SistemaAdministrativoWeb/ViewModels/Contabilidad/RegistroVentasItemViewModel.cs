namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class RegistroVentasItemViewModel
{
    public DateOnly FechaEmision { get; set; }
    public DateOnly FechaContabilizacion { get; set; }
    public string TipoComprobante { get; set; } = string.Empty;
    public string DescripcionTipoComprobante { get; set; } = string.Empty;
    public string Serie { get; set; } = string.Empty;
    public string Numero { get; set; } = string.Empty;
    public string CodigoPersona { get; set; } = string.Empty;
    public string NumeroDocumentoPersona { get; set; } = string.Empty;
    public string NombrePersona { get; set; } = string.Empty;
    public string CodigoMoneda { get; set; } = string.Empty;
    public decimal TipoCambio { get; set; }
    public decimal BaseImponible { get; set; }
    public decimal Descuento { get; set; }
    public decimal TotalExonerado { get; set; }
    public decimal TotalInafecto { get; set; }
    public decimal Igv { get; set; }
    public decimal Isc { get; set; }
    public decimal OtrosTributos { get; set; }
    public decimal Icbper { get; set; }
    public decimal Redondeo { get; set; }
    public decimal ImporteTotal { get; set; }
    public string Estado { get; set; } = string.Empty;
    public string Observacion { get; set; } = string.Empty;
}
