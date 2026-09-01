using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.ViewFeatures;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class ModuleSavePermissionFilter(
    IModulePermissionService modulePermissionService,
    ITempDataDictionaryFactory tempDataDictionaryFactory,
    string moduleCode,
    string identifierProperty) : IAsyncActionFilter
{
    public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        var isEdit = context.ActionArguments.Values
            .Where(value => value is not null)
            .Select(value => ResolvePropertyPath(value!, identifierProperty))
            .Any(HasIdentifier);

        var operation = isEdit
            ? ModulePermissionOperation.Edit
            : ModulePermissionOperation.Create;
        var access = await modulePermissionService.EvaluateAccessAsync(
            context.HttpContext.User,
            moduleCode,
            operation,
            context.HttpContext.RequestAborted);

        if (access.IsAllowed)
        {
            await next();
            return;
        }

        var tempData = tempDataDictionaryFactory.GetTempData(context.HttpContext);
        tempData["ErrorMessage"] = access.Message ?? "No tiene permiso para realizar esta operacion.";
        context.Result = access.Scope == ModuleScope.Company && !access.HasCompanyContext
            ? new RedirectToActionResult("Index", "EmpresaContexto", new { area = "" })
            : new RedirectToActionResult("Index", "Home", new { area = "" });
    }

    private static bool HasIdentifier(object? value)
    {
        return value switch
        {
            int integer => integer > 0,
            long integer => integer > 0,
            short integer => integer > 0,
            Guid guid => guid != Guid.Empty,
            string text => !string.IsNullOrWhiteSpace(text),
            _ => value is not null
        };
    }

    private static object? ResolvePropertyPath(object source, string propertyPath)
    {
        object? current = source;
        foreach (var segment in propertyPath.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (current is null)
            {
                return null;
            }

            current = current.GetType().GetProperty(segment)?.GetValue(current);
        }

        return current;
    }
}
