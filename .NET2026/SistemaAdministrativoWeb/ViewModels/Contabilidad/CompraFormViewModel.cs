using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CompraFormViewModel
{
    public int? IdCompra { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione un proveedor.")]
    public int? IdProveedor { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la configuracion contable.")]
    public int? IdConfiguracionContabilizacion { get; set; }

    [DataType(DataType.Date)]
    public DateOnly FechaEmision { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [DataType(DataType.Date)]
    public DateOnly FechaContabilizacion { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Required]
    [StringLength(3)]
    public string TipoComprobante { get; set; } = "01";

    [Required]
    [StringLength(10, MinimumLength = 1, ErrorMessage = "La serie debe tener entre 1 y 10 caracteres.")]
    [RegularExpression(@"^[A-Z0-9]{1,10}$", ErrorMessage = "La serie solo puede contener letras y numeros.")]
    public string Serie { get; set; } = string.Empty;

    [Required]
    [StringLength(20, ErrorMessage = "El numero debe tener como maximo 20 caracteres.")]
    [RegularExpression(@"^[A-Z0-9]{1,20}$", ErrorMessage = "El numero solo puede contener letras y numeros, con un maximo de 20 posiciones.")]
    public string Numero { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la moneda.")]
    public int? IdMoneda { get; set; }

    [Range(typeof(decimal), "0.000001", "999999999999")]
    public decimal TipoCambio { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal BaseImponible { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal TotalExonerado { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal TotalInafecto { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Icbper { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Igv { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Isc { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal OtrosTributos { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Redondeo { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ImporteTotal { get; set; }

    public bool ExoneracionRenta4ta { get; set; }

    [Range(typeof(decimal), "0", "100")]
    public decimal PorcentajeRetencion { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Retencion { get; set; }

    public bool TieneDetraccion { get; set; }

    public int? IdDetraccionSunat { get; set; }

    public string? DetraccionTexto { get; set; }

    [Range(typeof(decimal), "0", "100")]
    public decimal PorcentajeDetraccion { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ImporteDetraccion { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal SaldoPago { get; set; }

    public bool TienePercepcion { get; set; }

    public int? IdTipoPercepcion { get; set; }

    public string? TipoPercepcionTexto { get; set; }

    [Range(typeof(decimal), "0", "100")]
    public decimal PorcentajePercepcion { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal BasePercepcion { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ImportePercepcion { get; set; }

    [StringLength(500)]
    public string? Observacion { get; set; }

    public List<CompraDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Item = 1
        }
    ];
}
