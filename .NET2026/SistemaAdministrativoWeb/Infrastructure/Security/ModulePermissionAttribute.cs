using Microsoft.AspNetCore.Mvc;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true, Inherited = true)]
public sealed class ModulePermissionAttribute : TypeFilterAttribute
{
    public ModulePermissionAttribute(string moduleCode, ModulePermissionOperation operation = ModulePermissionOperation.Auto)
        : base(typeof(ModulePermissionFilter))
    {
        Arguments = [moduleCode, operation];
    }
}
