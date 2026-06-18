using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VentaDetalleFormViewModel
{
    [Range(1, short.MaxValue, ErrorMessage = "El item debe ser mayor a cero.")]
    public short Item { get; set; }

    [Required(ErrorMessage = "Ingrese la descripcion del concepto.")]
    [StringLength(250)]
    public string Descripcion { get; set; } = string.Empty;

    [Range(typeof(decimal), "0.0001", "999999999999")]
    public decimal Cantidad { get; set; } = 1m;

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ValorUnitario { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal ImporteBruto { get; set; }
}
