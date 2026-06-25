using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionNotaCreditoFormViewModel
{
    [Required(ErrorMessage = "Seleccione si la aplicacion es para cliente o proveedor.")]
    [StringLength(1)]
    public string TipoPersona { get; set; } = "C";

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la persona a trabajar.")]
    public int? IdPersona { get; set; }

    [StringLength(200)]
    public string PersonaTexto { get; set; } = string.Empty;

    [StringLength(20)]
    public string NumeroDocumentoPersona { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la fecha de aplicacion.")]
    [DataType(DataType.Date)]
    public DateOnly FechaAplicacion { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    public int? IdMoneda { get; set; }

    [StringLength(50)]
    public string MonedaTexto { get; set; } = string.Empty;

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese un tipo de cambio mayor a cero.")]
    public decimal TipoCambio { get; set; } = 1m;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione el comprobante pendiente.")]
    public int? IdRegistroComprobante { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la nota de credito.")]
    public int? IdRegistroNotaCredito { get; set; }

    [Range(typeof(decimal), "0.01", "999999999999", ErrorMessage = "Ingrese un importe aplicado mayor a cero.")]
    public decimal ImporteAplicado { get; set; }

    [Required(ErrorMessage = "Ingrese la glosa de la aplicacion.")]
    [StringLength(300)]
    public string Glosa { get; set; } = string.Empty;

    [StringLength(500)]
    public string? Observacion { get; set; }
}
