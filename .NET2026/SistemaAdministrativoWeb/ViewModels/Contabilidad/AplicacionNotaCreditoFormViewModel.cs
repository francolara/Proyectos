using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionNotaCreditoFormViewModel
{
    [Required(ErrorMessage = "Seleccione si la aplicacion es para cliente o proveedor.")]
    [StringLength(1, ErrorMessage = "El tipo de persona no puede exceder 1 caracter.")]
    public string TipoPersona { get; set; } = "C";

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la persona a trabajar.")]
    public int? IdPersona { get; set; }

    [StringLength(200, ErrorMessage = "El nombre de la persona no puede exceder 200 caracteres.")]
    public string PersonaTexto { get; set; } = string.Empty;

    [StringLength(20, ErrorMessage = "El numero de documento no puede exceder 20 caracteres.")]
    public string NumeroDocumentoPersona { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la fecha de aplicacion.")]
    [DataType(DataType.Date)]
    public DateOnly FechaAplicacion { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    public int? IdMoneda { get; set; }

    [StringLength(50, ErrorMessage = "La moneda no puede exceder 50 caracteres.")]
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
    [StringLength(300, ErrorMessage = "La glosa no puede exceder 300 caracteres.")]
    public string Glosa { get; set; } = string.Empty;

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; set; }
}
