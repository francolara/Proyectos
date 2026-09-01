using Microsoft.AspNetCore.Mvc;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
public sealed class ModuleSavePermissionAttribute : TypeFilterAttribute
{
    public ModuleSavePermissionAttribute(string moduleCode, string identifierProperty)
        : base(typeof(ModuleSavePermissionFilter))
    {
        Arguments = [moduleCode, identifierProperty];
    }
}
