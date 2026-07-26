using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class SubscriptionAccessMiddleware(RequestDelegate next)
{
    public async Task InvokeAsync(
        HttpContext context,
        ISubscriptionAccessService subscriptionAccessService)
    {
        if (!ShouldEvaluate(context))
        {
            await next(context);
            return;
        }

        var evaluation = await subscriptionAccessService.EvaluateAsync(
            context.User,
            context.RequestAborted);

        if (!evaluation.IsRestricted || IsAllowedEndpoint(context))
        {
            await next(context);
            return;
        }

        if (HttpMethods.IsGet(context.Request.Method)
            || HttpMethods.IsHead(context.Request.Method))
        {
            context.Response.Redirect("/MiSuscripcion");
            return;
        }

        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        await context.Response.WriteAsJsonAsync(
            new
            {
                mensaje = evaluation.Message,
                estado = evaluation.EffectiveStatus
            },
            context.RequestAborted);
    }

    private static bool ShouldEvaluate(HttpContext context)
    {
        if (context.User.Identity?.IsAuthenticated != true
            || context.User.IsInRole("SuperAdmin"))
        {
            return false;
        }

        var path = context.Request.Path;
        return !path.StartsWithSegments("/css", StringComparison.OrdinalIgnoreCase)
            && !path.StartsWithSegments("/js", StringComparison.OrdinalIgnoreCase)
            && !path.StartsWithSegments("/lib", StringComparison.OrdinalIgnoreCase)
            && !path.StartsWithSegments("/images", StringComparison.OrdinalIgnoreCase)
            && !path.StartsWithSegments("/favicon.ico", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAllowedEndpoint(HttpContext context)
    {
        if (context.GetEndpoint()?.Metadata.GetMetadata<AllowRestrictedSubscriptionAttribute>() is not null)
        {
            return true;
        }

        var path = context.Request.Path;
        return path.StartsWithSegments("/Identity/Account/Logout", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Identity/Account/VerificacionTemporal", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Identity/Account/CambiarContrasenaTemporal", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Home/Error", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Suscripcion/Renovar", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Suscripcion/Pago", StringComparison.OrdinalIgnoreCase);
    }
}
