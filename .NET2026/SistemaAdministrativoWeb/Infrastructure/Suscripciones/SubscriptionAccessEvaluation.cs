namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class SubscriptionAccessEvaluation
{
    public ContextoLoginUsuarioDto? LoginContext { get; init; }
    public bool IsRestricted { get; init; }
    public string EffectiveStatus { get; init; } = "SIN_CONTEXTO";
    public string StatusDisplay { get; init; } = "Sin contexto";
    public string PlanDisplay { get; init; } = "Sin plan";
    public DateOnly? ExpirationDate { get; init; }
    public DateOnly? GraceEndDate { get; init; }
    public int? DaysUntilExpiration { get; init; }
    public string Message { get; init; } = string.Empty;
}
