using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VentaDetalleFormViewModel
{
    [Range(1, short.MaxValue, ErrorMessage = "El item debe ser mayor a cero.")]
    public short Item { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la cuenta contable.")]
    public int? IdPlanCuenta { get; set; }

    public string CuentaTexto { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione el tipo de afectacion IGV.")]
    public int? IdTipoAfectacionIGV { get; set; } = 1;

    [Required(ErrorMessage = "Ingrese la descripcion del concepto.")]
    [StringLength(250, ErrorMessage = "La descripcion no puede exceder 250 caracteres.")]
    public string Descripcion { get; set; } = string.Empty;

    [Range(typeof(decimal), "0.0001", "999999999999")]
    public decimal Cantidad { get; set; } = 1m;

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ValorUnitario { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ImporteBruto { get; set; }
}
