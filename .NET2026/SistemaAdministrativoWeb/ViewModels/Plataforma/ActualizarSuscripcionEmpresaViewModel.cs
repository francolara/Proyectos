using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class ActualizarSuscripcionCuentaViewModel
{
    [Required]
    public int IdCuentaAdministradora { get; init; }

    [Required]
    [StringLength(50, ErrorMessage = "El tipo de plan no puede exceder 50 caracteres.")]
    public string TipoPlan { get; init; } = "TRIAL";

    [Required]
    [StringLength(20, ErrorMessage = "El estado de suscripcion no puede exceder 20 caracteres.")]
    public string EstadoSuscripcion { get; init; } = "TRIAL";

    public bool EsPrueba { get; init; }
    public DateOnly? FechaInicioPrueba { get; init; }
    public DateOnly? FechaFinPrueba { get; init; }
    public DateOnly? FechaInicioPlan { get; init; }
    public DateOnly? FechaFinPlan { get; init; }
    [StringLength(20, ErrorMessage = "El tipo de cobro no puede exceder 20 caracteres.")]
    public string? TipoCobro { get; init; }
    [Range(0, 60)]
    public int DiasGracia { get; init; } = 5;
    public int? EmpresasPermitidas { get; init; }
    public int? UsuariosPermitidos { get; init; }
    public bool Activo { get; init; }
    public bool EstadoCuenta { get; init; }

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; init; }
}
