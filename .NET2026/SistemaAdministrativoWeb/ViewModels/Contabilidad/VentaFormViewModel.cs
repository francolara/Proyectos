using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VentaFormViewModel
{
    public int? IdVenta { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione un cliente.")]
    public int? IdCliente { get; set; }

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
    [StringLength(10)]
    public string Serie { get; set; } = string.Empty;

    [Required]
    [StringLength(20)]
    public string Numero { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la moneda.")]
    public int? IdMoneda { get; set; }

    [Range(typeof(decimal), "0.000001", "999999999999")]
    public decimal TipoCambio { get; set; } = 1m;

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal BaseImponible { get; set; }

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

    [StringLength(500)]
    public string? Observacion { get; set; }

    public List<VentaDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Item = 1
        }
    ];
}
