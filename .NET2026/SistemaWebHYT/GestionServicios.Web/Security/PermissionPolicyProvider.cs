using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Options;

namespace GestionServicios.Web.Security;

public sealed class PermissionPolicyProvider(IOptions<AuthorizationOptions> options) : DefaultAuthorizationPolicyProvider(options)
{
    public override Task<AuthorizationPolicy?> GetPolicyAsync(string policyName)
        => policyName.StartsWith("perm:", StringComparison.Ordinal)
            ? Task.FromResult<AuthorizationPolicy?>(new AuthorizationPolicyBuilder().RequireAuthenticatedUser()
                .RequireClaim(AccessService.PermissionClaim, policyName[5..]).Build())
            : base.GetPolicyAsync(policyName);
}
