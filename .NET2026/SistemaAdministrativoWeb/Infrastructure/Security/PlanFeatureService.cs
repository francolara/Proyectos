using System.Security.Claims;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class PlanFeatureService(
    ISubscriptionAccessService subscriptionAccessService) : IPlanFeatureService
{
    public async Task<bool> IsEnabledAsync(
        ClaimsPrincipal principal,
        PlanFeature feature,
        CancellationToken cancellationToken = default)
    {
        if (principal.Identity?.IsAuthenticated != true)
        {
            return false;
        }

        if (principal.IsInRole("SuperAdmin"))
        {
            return true;
        }

        var access = await subscriptionAccessService.EvaluateAsync(principal, cancellationToken);
        if (access.IsRestricted || access.LoginContext is null)
        {
            return false;
        }

        return feature switch
        {
            PlanFeature.CpeValidation => HasCpeValidation(access.LoginContext),
            _ => false
        };
    }

    private static bool HasCpeValidation(ContextoLoginUsuarioDto context)
    {
        if (context.EsPrueba == true)
        {
            return false;
        }

        var plan = (context.TipoPlan ?? string.Empty).Trim().ToUpperInvariant();
        return plan is "PRO" or "CONTADOR";
    }
}
