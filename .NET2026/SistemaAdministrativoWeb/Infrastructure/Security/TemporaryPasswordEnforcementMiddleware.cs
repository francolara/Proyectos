using System.Security.Claims;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public sealed class TemporaryPasswordEnforcementMiddleware(RequestDelegate next)
{
    public async Task InvokeAsync(HttpContext context)
    {
        if (DebeAplicarRestriccion(context))
        {
            var userId = context.User.FindFirstValue(ClaimTypes.NameIdentifier);
            var verificado = !string.IsNullOrWhiteSpace(userId)
                && string.Equals(
                    context.Session.GetString(TemporaryPasswordFlowConstants.VerificationSessionKey),
                    userId,
                    StringComparison.Ordinal);

            var path = context.Request.Path;
            var enVerificacion = path.StartsWithSegments("/Identity/Account/VerificacionTemporal", StringComparison.OrdinalIgnoreCase);
            var enCambioClave = path.StartsWithSegments("/Identity/Account/CambiarContrasenaTemporal", StringComparison.OrdinalIgnoreCase);

            if (enVerificacion && verificado)
            {
                context.Response.Redirect("/Identity/Account/CambiarContrasenaTemporal");
                return;
            }

            if (enCambioClave && !verificado)
            {
                context.Response.Redirect("/Identity/Account/VerificacionTemporal");
                return;
            }

            if (!enVerificacion && !enCambioClave)
            {
                context.Response.Redirect(verificado
                    ? "/Identity/Account/CambiarContrasenaTemporal"
                    : "/Identity/Account/VerificacionTemporal");
                return;
            }
        }

        await next(context);
    }

    private static bool DebeAplicarRestriccion(HttpContext context)
    {
        if (context.User.Identity?.IsAuthenticated != true || context.User.IsInRole("SuperAdmin"))
        {
            return false;
        }

        if (!context.User.HasClaim(
                TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType,
                TemporaryPasswordFlowConstants.RequirePasswordChangeClaimValue))
        {
            return false;
        }

        var path = context.Request.Path;
        if (path.HasValue && (Path.HasExtension(path) || path.StartsWithSegments("/lib", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        if (path.StartsWithSegments("/Identity/Account/Logout", StringComparison.OrdinalIgnoreCase)
            || path.StartsWithSegments("/Home/Error", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }
}
