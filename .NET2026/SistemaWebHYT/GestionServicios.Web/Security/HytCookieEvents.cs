using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Security;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;

namespace GestionServicios.Web.Security;

public sealed class HytCookieEvents(IAccessRepository repository, TimeProvider clock, ILogger<HytCookieEvents> logger) : CookieAuthenticationEvents
{
    public override async Task ValidatePrincipal(CookieValidatePrincipalContext context)
    {
        try
        {
            var id = context.Principal?.FindFirstValue(ClaimTypes.NameIdentifier);
            if (id is null || !context.Properties.Items.TryGetValue("hyt:created", out var created) ||
                !long.TryParse(created, NumberStyles.None, CultureInfo.InvariantCulture, out var ticks) ||
                clock.GetUtcNow().UtcTicks - ticks > TimeSpan.FromHours(8).Ticks)
            { context.RejectPrincipal(); await context.HttpContext.SignOutAsync("HytCookie"); return; }
            var account = await repository.FindByIdAsync(id, context.HttpContext.RequestAborted);
            if (account is null || !account.Active || account.SecurityStamp != context.Principal?.FindFirstValue(AccessService.StampClaim) ||
                account.LockoutEnd > clock.GetUtcNow())
            { context.RejectPrincipal(); await context.HttpContext.SignOutAsync("HytCookie"); return; }
            // Releer permisos impide que una cookie conserve autorizaciones retiradas.
            context.ReplacePrincipal(AccessService.Principal(account, ticks + TimeSpan.FromHours(8).Ticks));
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            logger.LogWarning("No se pudo validar la sesión HYT. Acceso denegado.");
            context.RejectPrincipal();
        }
        if (context.Principal?.Identity?.IsAuthenticated != true)
            await context.HttpContext.SignOutAsync("HytCookie");
    }
}
