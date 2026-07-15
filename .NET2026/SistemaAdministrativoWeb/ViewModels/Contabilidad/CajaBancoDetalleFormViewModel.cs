using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoDetalleFormViewModel
{
    public int? IdMovimientoBancoDetalle { get; set; }

    [Range(1, short.MaxValue, ErrorMessage = "El item debe ser mayor a cero.")]
    public short Item { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una cuenta contable.")]
    public int? IdPlanCuenta { get; set; }

    public string CuentaTexto { get; set; } = string.Empty;
    public bool RequiereCentroCostoCuenta { get; set; }

    public int? IdPersona { get; set; }
    public string PersonaTexto { get; set; } = string.Empty;

    [StringLength(3, ErrorMessage = "El modulo del comprobante no puede exceder 3 caracteres.")]
    public string? ModuloOperacionComprobante { get; set; }

    public int? IdRegistroComprobante { get; set; }

    [StringLength(300, ErrorMessage = "La glosa de la linea no puede exceder 300 caracteres.")]
    public string? GlosaDetalle { get; set; }

    [StringLength(20, ErrorMessage = "El codigo del centro de costo no puede exceder 20 caracteres.")]
    public string? CodigoCentroCosto { get; set; }

    [StringLength(20, ErrorMessage = "El numero de documento no puede exceder 20 caracteres.")]
    public string? NumeroDocumento { get; set; }

    [StringLength(150, ErrorMessage = "El tipo de documento no puede exceder 150 caracteres.")]
    public string? TipoDocumento { get; set; }

    [StringLength(10, ErrorMessage = "La serie no puede exceder 10 caracteres.")]
    public string? Serie { get; set; }

    [StringLength(100, ErrorMessage = "La referencia de la linea no puede exceder 100 caracteres.")]
    public string? ReferenciaLinea { get; set; }

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese un tipo de cambio mayor a cero en la linea.")]
    public decimal TipoCambioLinea { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Debe { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Haber { get; set; }
}
