using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaFormViewModel
{
    public int? IdPlanCuenta { get; set; }

    [Display(Name = "Cuenta padre")]
    public int? IdPlanCuentaPadre { get; set; }

    [Required(ErrorMessage = "Ingrese el codigo de cuenta.")]
    [StringLength(20)]
    public string CodigoCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Ingrese el nombre de la cuenta.")]
    [StringLength(200)]
    public string NombreCuenta { get; set; } = string.Empty;

    [Required(ErrorMessage = "Seleccione la naturaleza.")]
    [StringLength(1)]
    public string NaturalezaSaldo { get; set; } = "D";

    public bool AceptaMovimiento { get; set; }
    public bool RequiereCentroCosto { get; set; }
    public bool Estado { get; set; } = true;
}
