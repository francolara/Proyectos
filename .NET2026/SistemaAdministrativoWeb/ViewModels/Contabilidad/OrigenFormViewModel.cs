using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class OrigenFormViewModel
{
    public int? IdOrigen { get; set; }

    [Required(ErrorMessage = "Ingrese el codigo del origen.")]
    [StringLength(10, ErrorMessage = "El codigo del origen no puede exceder 10 caracteres.")]
    public string CodigoOrigen { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre del origen.")]
    [StringLength(150, ErrorMessage = "El nombre del origen no puede exceder 150 caracteres.")]
    public string NombreOrigen { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el modulo de origen.")]
    [StringLength(50, ErrorMessage = "El modulo de origen no puede exceder 50 caracteres.")]
    public string ModuloOrigen { get; set; } = string.Empty;

    public bool PermiteRegistroManual { get; set; } = true;
    public bool Estado { get; set; } = true;
}
