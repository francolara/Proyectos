using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public interface IModulePermissionService
{
    Task<bool> CanAccessModuleAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        CancellationToken cancellationToken = default);

    Task<bool> CanAccessModuleAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        ModulePermissionOperation operation,
        CancellationToken cancellationToken = default);

    Task<ModuleAccessResult> EvaluateAccessAsync(
        ClaimsPrincipal principal,
        string moduleCode,
        ModulePermissionOperation operation,
        CancellationToken cancellationToken = default);
}
