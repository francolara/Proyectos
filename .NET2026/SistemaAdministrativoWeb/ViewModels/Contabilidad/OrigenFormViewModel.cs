using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class OrigenFormViewModel
{
    public int? IdOrigen { get; set; }

    [Required(ErrorMessage = "Ingrese el codigo del origen.")]
    [StringLength(10)]
    public string CodigoOrigen { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre del origen.")]
    [StringLength(150)]
    public string NombreOrigen { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el modulo de origen.")]
    [StringLength(50)]
    public string ModuloOrigen { get; set; } = string.Empty;

    public bool PermiteRegistroManual { get; set; } = true;
    public bool Estado { get; set; } = true;
}
