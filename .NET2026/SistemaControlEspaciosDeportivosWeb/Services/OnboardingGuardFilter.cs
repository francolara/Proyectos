using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class OnboardingGuardFilter(ISportCenterStoredProcedureService spService) : IAsyncActionFilter
{
    private static readonly HashSet<string> ControladoresProtegidos = new(StringComparer.OrdinalIgnoreCase)
    {
        "Reservas",
        "Pagos",
        "Comprobantes",
        "Reportes",
        "Promociones",
        "Cupones"
    };

    public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
    {
        var user = context.HttpContext.User;
        if (user?.Identity?.IsAuthenticated != true)
        {
            await next();
            return;
        }

        var controller = Convert.ToString(context.RouteData.Values["controller"]) ?? string.Empty;
        if (!ControladoresProtegidos.Contains(controller))
        {
            await next();
            return;
        }

        var negocioId = await ResolverNegocioIdAsync(context);
        if (!negocioId.HasValue)
        {
            await next();
            return;
        }

        var checklist = await spService.OnboardingChecklistValidarAsync(negocioId.Value);
        if (checklist.ChecklistCompleto)
        {
            await next();
            return;
        }

        context.Result = new RedirectToActionResult("Index", "Onboarding", new { negocioId = negocioId.Value });
    }

    private async Task<int?> ResolverNegocioIdAsync(ActionExecutingContext context)
    {
        if (TryReadInt(context.ActionArguments, "negocioId", out var negocioIdArg) && negocioIdArg > 0)
            return negocioIdArg;
        if (TryReadInt(context.ActionArguments, "NegocioId", out var negocioIdArg2) && negocioIdArg2 > 0)
            return negocioIdArg2;

        if (TryReadInt(context.HttpContext.Request.Query, "negocioId", out var negocioIdQuery) && negocioIdQuery > 0)
            return negocioIdQuery;

        var routeNegocio = Convert.ToString(context.RouteData.Values["negocioId"]);
        if (int.TryParse(routeNegocio, out var negocioIdRoute) && negocioIdRoute > 0)
            return negocioIdRoute;

        var usuarioId = context.HttpContext.User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId))
            return null;

        var membresias = await spService.PanelListarNegociosUsuarioAsync(usuarioId);
        return membresias.FirstOrDefault()?.NegocioId;
    }

    private static bool TryReadInt(IDictionary<string, object?> args, string key, out int value)
    {
        value = 0;
        if (!args.TryGetValue(key, out var raw) || raw is null)
            return false;

        return raw switch
        {
            int i => (value = i) > 0,
            string s when int.TryParse(s, out var p) => (value = p) > 0,
            _ => int.TryParse(Convert.ToString(raw), out var p2) && (value = p2) > 0
        };
    }

    private static bool TryReadInt(IQueryCollection query, string key, out int value)
    {
        value = 0;
        if (!query.TryGetValue(key, out var raw))
            return false;

        return int.TryParse(raw.FirstOrDefault(), out value);
    }
}
