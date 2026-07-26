namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class ActivarContratoCuentaRequest
{
    public int IdCuentaAdministradora { get; init; }
    public string TipoPlan { get; init; } = "BASICO";
    public string TipoCobro { get; init; } = "MENSUAL";
    public DateOnly FechaInicioPlan { get; init; }
    public DateOnly FechaFinPlan { get; init; }
    public int DiasGracia { get; init; } = 5;
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
}
