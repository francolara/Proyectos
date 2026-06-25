using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CentroCostoFormViewModel
{
    public int? IdCentroCosto { get; set; }

    [Required(ErrorMessage = "Ingrese el codigo del centro de costo.")]
    [StringLength(20, ErrorMessage = "El codigo no puede exceder 20 caracteres.")]
    public string CodigoCentroCosto { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre del centro de costo.")]
    [StringLength(150, ErrorMessage = "El nombre no puede exceder 150 caracteres.")]
    public string NombreCentroCosto { get; set; } = string.Empty;

    public bool Estado { get; set; } = true;
}
