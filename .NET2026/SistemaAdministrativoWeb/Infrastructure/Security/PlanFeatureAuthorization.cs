using Microsoft.AspNetCore.Authorization;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public static class PlanFeaturePolicies
{
    public const string CpeValidation = "PlanFeature:CpeValidation";
}

public sealed class PlanFeatureRequirement(PlanFeature feature) : IAuthorizationRequirement
{
    public PlanFeature Feature { get; } = feature;
}

public sealed class PlanFeatureAuthorizationHandler(
    IPlanFeatureService planFeatureService,
    IHttpContextAccessor httpContextAccessor) : AuthorizationHandler<PlanFeatureRequirement>
{
    protected override async Task HandleRequirementAsync(
        AuthorizationHandlerContext context,
        PlanFeatureRequirement requirement)
    {
        var cancellationToken = httpContextAccessor.HttpContext?.RequestAborted ?? CancellationToken.None;
        if (await planFeatureService.IsEnabledAsync(context.User, requirement.Feature, cancellationToken))
        {
            context.Succeed(requirement);
        }
    }
}
