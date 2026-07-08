namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class RegistroComprasItemViewModel
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
    public decimal BaseImponibleGravada { get; set; }
    public decimal IgvGravado { get; set; }
    public decimal BaseImponibleGasto { get; set; }
    public decimal IgvGasto { get; set; }
    public decimal BaseImponibleSinCredito { get; set; }
    public decimal IgvSinCredito { get; set; }
    public decimal TotalExonerado { get; set; }
    public decimal TotalInafecto { get; set; }
    public decimal OtrosTributos { get; set; }
    public decimal Icbper { get; set; }
    public decimal Retencion { get; set; }
    public decimal ImporteDetraccion { get; set; }
    public decimal ImportePercepcion { get; set; }
    public decimal ImporteTotal { get; set; }
    public string Estado { get; set; } = string.Empty;
    public string Observacion { get; set; } = string.Empty;
}
