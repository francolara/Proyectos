using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class ActivarContratoCuentaViewModel
{
    [Required]
    public int IdCuentaAdministradora { get; init; }

    [Required]
    [StringLength(20)]
    public string TipoCobro { get; init; } = "MENSUAL";

    public DateOnly FechaInicioPlan { get; init; } = DateOnly.FromDateTime(DateTime.Today);

    [Range(0, 60)]
    public int DiasGracia { get; init; } = 5;

    [StringLength(500)]
    public string? Observacion { get; init; }
}
