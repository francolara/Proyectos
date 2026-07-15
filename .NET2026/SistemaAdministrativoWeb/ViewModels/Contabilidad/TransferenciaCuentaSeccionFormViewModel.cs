using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TransferenciaCuentaSeccionFormViewModel
{
    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una cuenta corriente.")]
    public int? IdBancoConfiguracionEmpresa { get; set; }

    [Required(ErrorMessage = "Seleccione la operacion bancaria.")]
    [StringLength(2, ErrorMessage = "La operacion bancaria no puede exceder 2 caracteres.")]
    public string IdOpeBancaria { get; set; } = string.Empty;

    public string CuentaCorrienteTexto { get; set; } = string.Empty;
    public string MonedaTexto { get; set; } = string.Empty;
    public string TipoOperacionTexto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese la fecha de emision.")]
    [DataType(DataType.Date)]
    public DateOnly FechaEmision { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [Range(typeof(decimal), "0.000001", "999999999999", ErrorMessage = "Ingrese un tipo de cambio mayor a cero.")]
    public decimal TipoCambio { get; set; } = 1m;

    [StringLength(20, ErrorMessage = "El numero de operacion no puede exceder 20 caracteres.")]
    public string? NumeroOperacion { get; set; }

    [Range(typeof(decimal), "0", "999999999999", ErrorMessage = "Ingrese un monto valido.")]
    public decimal Importe { get; set; }

    [Required(ErrorMessage = "Ingrese la glosa del movimiento.")]
    [StringLength(300, ErrorMessage = "La glosa no puede exceder 300 caracteres.")]
    public string Glosa { get; set; } = string.Empty;

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; set; }
}
