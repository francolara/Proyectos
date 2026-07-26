using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public interface ISubscriptionAccessService
{
    Task<SubscriptionAccessEvaluation> EvaluateAsync(
        ClaimsPrincipal principal,
        CancellationToken cancellationToken = default);
}
