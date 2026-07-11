using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.ViewFeatures;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class ModulePermissionFilter(
    IModulePermissionService modulePermissionService,
    ITempDataDictionaryFactory tempDataDictionaryFactory,
    string moduleCode,
    ModulePermissionOperation operation) : IAsyncAuthorizationFilter
{
    public async Task OnAuthorizationAsync(AuthorizationFilterContext context)
    {
        if (context.HttpContext.User.Identity?.IsAuthenticated != true)
        {
            context.Result = new ChallengeResult();
            return;
        }

        var resolvedOperation = operation == ModulePermissionOperation.Auto
            ? ResolveOperation(context)
            : operation;

        var access = await modulePermissionService.EvaluateAccessAsync(
            context.HttpContext.User,
            moduleCode,
            resolvedOperation,
            context.HttpContext.RequestAborted);

        if (access.IsAllowed)
        {
            return;
        }

        var tempData = tempDataDictionaryFactory.GetTempData(context.HttpContext);
        tempData["ErrorMessage"] = access.Message ?? "No tiene permiso para acceder a esta opcion.";

        context.Result = access.Scope == ModuleScope.Company && !access.HasCompanyContext
            ? new RedirectToActionResult("Index", "EmpresaContexto", new { area = "" })
            : new RedirectToActionResult("Index", "Home", new { area = "" });
    }

    private static ModulePermissionOperation ResolveOperation(AuthorizationFilterContext context)
    {
        var actionName = context.ActionDescriptor.RouteValues.TryGetValue("action", out var routeAction)
            ? routeAction ?? string.Empty
            : string.Empty;
        var httpMethod = context.HttpContext.Request.Method.ToUpperInvariant();

        if (actionName.StartsWith("Eliminar", StringComparison.OrdinalIgnoreCase))
        {
            return ModulePermissionOperation.Delete;
        }

        if (actionName.StartsWith("Editar", StringComparison.OrdinalIgnoreCase)
            || actionName.StartsWith("Validar", StringComparison.OrdinalIgnoreCase)
            || actionName.StartsWith("Actualizar", StringComparison.OrdinalIgnoreCase))
        {
            return ModulePermissionOperation.Edit;
        }

        if (actionName.StartsWith("Guardar", StringComparison.OrdinalIgnoreCase)
            || actionName.StartsWith("Previsualizar", StringComparison.OrdinalIgnoreCase))
        {
            return ModulePermissionOperation.CreateOrEdit;
        }

        if (actionName.StartsWith("Registrar", StringComparison.OrdinalIgnoreCase)
            || actionName.StartsWith("Crear", StringComparison.OrdinalIgnoreCase)
            || actionName.StartsWith("CargaMasiva", StringComparison.OrdinalIgnoreCase))
        {
            return ModulePermissionOperation.Create;
        }

        if (httpMethod == "POST" && !actionName.StartsWith("Buscar", StringComparison.OrdinalIgnoreCase))
        {
            return ModulePermissionOperation.CreateOrEdit;
        }

        return ModulePermissionOperation.View;
    }
}
