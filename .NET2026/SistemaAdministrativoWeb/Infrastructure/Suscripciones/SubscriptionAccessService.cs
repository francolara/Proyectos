using System.Globalization;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class SubscriptionAccessService(
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ILogger<SubscriptionAccessService> logger) : ISubscriptionAccessService
{
    private string? evaluatedUserId;
    private SubscriptionAccessEvaluation? cachedEvaluation;

    public async Task<SubscriptionAccessEvaluation> EvaluateAsync(
        ClaimsPrincipal principal,
        CancellationToken cancellationToken = default)
    {
        if (principal.Identity?.IsAuthenticated != true)
        {
            return CreateUnrestricted("SIN_AUTENTICAR", "Sin autenticar");
        }

        if (principal.IsInRole("SuperAdmin"))
        {
            return CreateUnrestricted("SUPERADMIN", "SuperAdmin");
        }

        var userId = principal.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(userId))
        {
            return CreateUnrestricted("SIN_USUARIO", "Sin usuario");
        }

        if (cachedEvaluation is not null
            && string.Equals(evaluatedUserId, userId, StringComparison.Ordinal))
        {
            return cachedEvaluation;
        }

        evaluatedUserId = userId;
        var context = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(userId, cancellationToken);
        cachedEvaluation = EvaluateContext(context, DateOnly.FromDateTime(DateTime.Today));

        if (ShouldPersistExpiredSuspension(context, cachedEvaluation))
        {
            try
            {
                await cuentaAdministradoraRepository.SincronizarVencimientoSuscripcionCuentaAsync(
                    context!.IdCuentaAdministradora!.Value,
                    userId,
                    cancellationToken);
            }
            catch (Exception exception)
            {
                logger.LogWarning(
                    exception,
                    "No se pudo persistir la suspension por vencimiento de la cuenta administradora {IdCuentaAdministradora}.",
                    context!.IdCuentaAdministradora);
            }
        }

        return cachedEvaluation;
    }

    private static bool ShouldPersistExpiredSuspension(
        ContextoLoginUsuarioDto? context,
        SubscriptionAccessEvaluation evaluation)
    {
        if (context?.IdCuentaAdministradora is null
            || !evaluation.IsRestricted
            || !string.Equals(evaluation.EffectiveStatus, "VENCIDO", StringComparison.Ordinal))
        {
            return false;
        }

        var rawStatus = (context.EstadoSuscripcion ?? string.Empty).Trim().ToUpperInvariant();
        return rawStatus is not "SUSPENDIDO" and not "BAJA";
    }

    private static SubscriptionAccessEvaluation EvaluateContext(
        ContextoLoginUsuarioDto? context,
        DateOnly today)
    {
        if (context is null || !context.TieneAcceso || !context.IdCuentaAdministradora.HasValue)
        {
            return CreateUnrestricted("SIN_CUENTA", "Sin cuenta", context);
        }

        var planDisplay = ResolvePlanDisplay(context);
        var rawStatus = (context.EstadoSuscripcion ?? string.Empty).Trim().ToUpperInvariant();

        if (context.IdCuentaAdministradoraSuscripcion is null)
        {
            return CreateRestricted(
                context,
                "SIN_SUSCRIPCION",
                "Sin suscripcion",
                planDisplay,
                null,
                null,
                null,
                "La cuenta no tiene una suscripcion configurada. El acceso operativo esta restringido.");
        }

        if (string.Equals(rawStatus, "BAJA", StringComparison.Ordinal))
        {
            return CreateRestricted(
                context,
                "BAJA",
                "Baja",
                planDisplay,
                ResolveExpirationDate(context),
                context.FechaFinGracia,
                CalculateDays(ResolveExpirationDate(context), today),
                "La suscripcion se encuentra de baja. El acceso operativo esta restringido.");
        }

        if (context.ActivoSuscripcion == false
            || string.Equals(rawStatus, "SUSPENDIDO", StringComparison.Ordinal))
        {
            return CreateRestricted(
                context,
                "SUSPENDIDO",
                "Suspendido",
                planDisplay,
                ResolveExpirationDate(context),
                context.FechaFinGracia,
                CalculateDays(ResolveExpirationDate(context), today),
                "La suscripcion esta suspendida. El acceso operativo esta restringido.");
        }

        if (IsTrial(context))
        {
            var trialEnd = context.FechaFinPrueba;
            if (!trialEnd.HasValue)
            {
                return CreateRestricted(
                    context,
                    "SIN_VIGENCIA",
                    "Sin vigencia",
                    planDisplay,
                    null,
                    null,
                    null,
                    "El periodo de prueba no tiene una fecha de vigencia configurada. El acceso operativo esta restringido.");
            }

            var trialDays = CalculateDays(trialEnd, today);
            if (today > trialEnd.Value)
            {
                return CreateRestricted(
                    context,
                    "VENCIDO",
                    "Vencido",
                    planDisplay,
                    trialEnd,
                    null,
                    trialDays,
                    $"El periodo de prueba vencio el {trialEnd:dd/MM/yyyy}. El acceso operativo esta restringido.");
            }

            return new SubscriptionAccessEvaluation
            {
                LoginContext = context,
                IsRestricted = false,
                EffectiveStatus = "EN_PRUEBA",
                StatusDisplay = "En prueba",
                PlanDisplay = planDisplay,
                ExpirationDate = trialEnd,
                DaysUntilExpiration = trialDays,
                Message = $"Periodo de prueba vigente hasta el {trialEnd:dd/MM/yyyy}."
            };
        }

        var planEnd = context.FechaFinPlan;
        if (!planEnd.HasValue)
        {
            return CreateRestricted(
                context,
                "SIN_VIGENCIA",
                "Sin vigencia",
                planDisplay,
                null,
                null,
                null,
                "El plan no tiene una fecha de vigencia configurada. El acceso operativo esta restringido.");
        }

        var planDays = CalculateDays(planEnd, today);
        if (today <= planEnd.Value)
        {
            return new SubscriptionAccessEvaluation
            {
                LoginContext = context,
                IsRestricted = false,
                EffectiveStatus = "ACTIVO",
                StatusDisplay = "Activo",
                PlanDisplay = planDisplay,
                ExpirationDate = planEnd,
                GraceEndDate = ResolveGraceEndDate(context, planEnd.Value),
                DaysUntilExpiration = planDays,
                Message = $"Plan vigente hasta el {planEnd:dd/MM/yyyy}."
            };
        }

        var graceEnd = ResolveGraceEndDate(context, planEnd.Value);
        if (graceEnd.HasValue && today <= graceEnd.Value)
        {
            return new SubscriptionAccessEvaluation
            {
                LoginContext = context,
                IsRestricted = false,
                EffectiveStatus = "EN_GRACIA",
                StatusDisplay = "En gracia",
                PlanDisplay = planDisplay,
                ExpirationDate = planEnd,
                GraceEndDate = graceEnd,
                DaysUntilExpiration = planDays,
                Message = $"El plan vencio el {planEnd:dd/MM/yyyy} y se encuentra en gracia hasta el {graceEnd:dd/MM/yyyy}."
            };
        }

        return CreateRestricted(
            context,
            "VENCIDO",
            "Vencido",
            planDisplay,
            planEnd,
            graceEnd,
            planDays,
            graceEnd.HasValue
                ? $"El periodo de gracia finalizo el {graceEnd:dd/MM/yyyy}. El acceso operativo esta restringido."
                : $"El plan vencio el {planEnd:dd/MM/yyyy}. El acceso operativo esta restringido.");
    }

    private static SubscriptionAccessEvaluation CreateUnrestricted(
        string status,
        string statusDisplay,
        ContextoLoginUsuarioDto? context = null)
    {
        return new SubscriptionAccessEvaluation
        {
            LoginContext = context,
            IsRestricted = false,
            EffectiveStatus = status,
            StatusDisplay = statusDisplay,
            PlanDisplay = ResolvePlanDisplay(context)
        };
    }

    private static SubscriptionAccessEvaluation CreateRestricted(
        ContextoLoginUsuarioDto context,
        string status,
        string statusDisplay,
        string planDisplay,
        DateOnly? expirationDate,
        DateOnly? graceEndDate,
        int? daysUntilExpiration,
        string message)
    {
        return new SubscriptionAccessEvaluation
        {
            LoginContext = context,
            IsRestricted = true,
            EffectiveStatus = status,
            StatusDisplay = statusDisplay,
            PlanDisplay = planDisplay,
            ExpirationDate = expirationDate,
            GraceEndDate = graceEndDate,
            DaysUntilExpiration = daysUntilExpiration,
            Message = message
        };
    }

    private static bool IsTrial(ContextoLoginUsuarioDto context)
    {
        var plan = (context.TipoPlan ?? string.Empty).Trim();
        var status = (context.EstadoSuscripcion ?? string.Empty).Trim();
        return context.EsPrueba == true
            || string.Equals(plan, "TRIAL", StringComparison.OrdinalIgnoreCase)
            || string.Equals(plan, "GRATIS", StringComparison.OrdinalIgnoreCase)
            || string.Equals(status, "TRIAL", StringComparison.OrdinalIgnoreCase);
    }

    private static string ResolvePlanDisplay(ContextoLoginUsuarioDto? context)
    {
        if (context is null)
        {
            return "Sin plan";
        }

        var plan = (context.TipoPlan ?? string.Empty).Trim().ToUpperInvariant();
        if (context.EsPrueba == true || plan is "TRIAL" or "GRATIS")
        {
            return "Prueba";
        }

        return plan switch
        {
            "BASICO" => "Emprendedor",
            "EMPRENDEDOR" => "Emprendedor",
            "PRO" => "Contador",
            "CONTADOR" => "Contador",
            "" => "Sin plan",
            _ => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(plan.ToLowerInvariant())
        };
    }

    private static DateOnly? ResolveExpirationDate(ContextoLoginUsuarioDto context)
        => IsTrial(context) ? context.FechaFinPrueba : context.FechaFinPlan;

    private static DateOnly? ResolveGraceEndDate(ContextoLoginUsuarioDto context, DateOnly planEnd)
    {
        if (context.FechaFinGracia.HasValue)
        {
            return context.FechaFinGracia;
        }

        var graceDays = context.DiasGracia.GetValueOrDefault();
        return graceDays > 0
            ? planEnd.AddDays(graceDays)
            : null;
    }

    private static int? CalculateDays(DateOnly? expirationDate, DateOnly today)
        => expirationDate?.DayNumber - today.DayNumber;
}
