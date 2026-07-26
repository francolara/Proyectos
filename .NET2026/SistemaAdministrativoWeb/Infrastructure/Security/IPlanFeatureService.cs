using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public interface IPlanFeatureService
{
    Task<bool> IsEnabledAsync(
        ClaimsPrincipal principal,
        PlanFeature feature,
        CancellationToken cancellationToken = default);
}
