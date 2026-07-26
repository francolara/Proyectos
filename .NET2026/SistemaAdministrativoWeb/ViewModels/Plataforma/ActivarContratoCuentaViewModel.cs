using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class ActivarContratoCuentaViewModel
{
    [Required]
    public int IdCuentaAdministradora { get; init; }

    [Required]
    [RegularExpression("^(BASICO|PRO)$", ErrorMessage = "Seleccione el plan Emprendedor o Contador.")]
    public string TipoPlan { get; init; } = "BASICO";

    [Required]
    [StringLength(20, ErrorMessage = "El tipo de cobro no puede exceder 20 caracteres.")]
    public string TipoCobro { get; init; } = "MENSUAL";

    public DateOnly FechaInicioPlan { get; init; } = DateOnly.FromDateTime(DateTime.Today);

    [Range(0, 60)]
    public int DiasGracia { get; init; } = 5;

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; init; }
}
