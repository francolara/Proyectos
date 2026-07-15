using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoFormViewModel
{
    public int? IdMovimientoBanco { get; set; }
    public int? NumeroMovimiento { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una cuenta corriente.")]
    public int? IdBancoConfiguracionEmpresa { get; set; }

    public string CuentaCorrienteTexto { get; set; } = string.Empty;
    public string BancoTexto { get; set; } = string.Empty;
    public string MonedaTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Seleccione si el movimiento es ingreso o egreso.")]
    [StringLength(1, ErrorMessage = "El tipo de movimiento no puede exceder 1 caracter.")]
    public string TipoMovimiento { get; set; } = "I";

    [Required(ErrorMessage = "Seleccione la operacion bancaria.")]
    [StringLength(2, ErrorMessage = "La operacion bancaria no puede exceder 2 caracteres.")]
    public string IdOpeBancaria { get; set; } = string.Empty;

    public string TipoOperacionTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la fecha de emision.")]
    [DataType(DataType.Date)]
    public DateOnly FechaEmision { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese un tipo de cambio mayor a cero.")]
    public decimal TipoCambio { get; set; }

    public int? IdPersona { get; set; }
    public string PersonaTexto { get; set; } = string.Empty;
    public string NumeroDocumentoPersona { get; set; } = string.Empty;

    [StringLength(20, ErrorMessage = "El numero de documento no puede exceder 20 caracteres.")]
    public string? NumeroDocumento { get; set; }

    [Required(ErrorMessage = "Ingrese la glosa del movimiento.")]
    [StringLength(300, ErrorMessage = "La glosa no puede exceder 300 caracteres.")]
    public string Glosa { get; set; } = string.Empty;

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; set; }

    [Range(typeof(decimal), "0.01", "999999999999", ErrorMessage = "Ingrese un importe total mayor a cero.")]
    public decimal ImporteTotal { get; set; }

    public List<CajaBancoDetalleFormViewModel> Detalles { get; set; } = [];
}
