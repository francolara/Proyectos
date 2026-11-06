using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Security;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.RateLimiting;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Acceso y salida MVC con antifalsificación, cookies, Turnstile validado en servidor y auditoría.
// Firma: FRANCO LARA - 05/11/2026 | Respaldo manual validado en servidor ante falta o fallo de Turnstile, sin omitir contraseña ni límites.
[Authorize]
public sealed class AccessController(AccessService access, IAccessRepository repository, ITurnstileVerifier turnstile, ManualCaptchaService captcha, TimeProvider clock, ILogger<AccessController> logger) : Controller
{
    [AllowAnonymous, HttpGet("/Acceso"), EnableRateLimiting("access")]
    public IActionResult Login(string? returnUrl) => User.Identity?.IsAuthenticated == true ? LocalRedirect("/") : LoginView(new AccessViewModel { ReturnUrl = returnUrl });
    [AllowAnonymous, HttpPost("/Acceso"), EnableRateLimiting("access")]
    public async Task<IActionResult> Login(AccessViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid) return LoginView(model);
        try
        {
            if(turnstile.Required && (model.UseManualCaptcha ? !captcha.Validate(HttpContext,model.CaptchaId,model.CaptchaAnswer) : !await turnstile.VerifyAsync(Request.Form["cf-turnstile-response"].ToString(),cancellationToken)))
            {
                ModelState.AddModelError("",model.UseManualCaptcha?"El captcha es incorrecto o venció. Escribe el nuevo código.":"No se pudo verificar el acceso automáticamente. Usa el captcha manual de respaldo.");
                model.UseManualCaptcha=true;
                return LoginView(model);
            }
            var result = await access.SignInAsync(model.UserName, model.Password, cancellationToken);
            if (result.Succeeded && result.Account is not null)
            {
                var properties = new AuthenticationProperties { IsPersistent = false, AllowRefresh = true };
                properties.Items["hyt:created"] = clock.GetUtcNow().UtcTicks.ToString(CultureInfo.InvariantCulture);
                await HttpContext.SignInAsync("HytCookie", AccessService.Principal(result.Account, clock.GetUtcNow().AddHours(8).UtcTicks), properties);
                return LocalRedirect(Url.IsLocalUrl(model.ReturnUrl) ? model.ReturnUrl! : "/");
            }
            ModelState.AddModelError("", "No se pudo iniciar sesión. Revisa tus datos o consulta con tu administrador.");
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        {
            logger.LogWarning("El servicio de acceso HYT no está disponible.");
            ModelState.AddModelError("", "No pudimos conectar con el servicio de acceso. Inténtalo nuevamente.");
        }
        model.Password = ""; ModelState.Remove(nameof(model.Password));
        return LoginView(model);
    }
    [AllowAnonymous, HttpGet("/Acceso/Captcha/Nuevo"), EnableRateLimiting("captcha")]
    public IActionResult NewCaptcha()
    {
        var id=captcha.Create(HttpContext);
        Response.Headers.CacheControl="no-store";
        return Json(new{id,imageUrl=Url.Action(nameof(CaptchaImage),new{id})});
    }
    [AllowAnonymous, HttpGet("/Acceso/Captcha/{id:guid}"), EnableRateLimiting("captcha")]
    public IActionResult CaptchaImage(Guid id)
    {
        Response.Headers.CacheControl="no-store";
        Response.Headers.ContentSecurityPolicy="default-src 'none'; style-src 'none'; sandbox";
        var svg=captcha.Image(HttpContext,id);
        return svg is null?NotFound():Content(svg,"image/svg+xml; charset=utf-8");
    }
    [HttpGet("/Salir")]
    public IActionResult Logout() => View();
    [HttpPost("/Salir"), ActionName(nameof(Logout))]
    public async Task<IActionResult> LogoutConfirmed(CancellationToken cancellationToken)
    {
        try
        {
            var id = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (id is not null) await repository.RecordLogoutAsync(id, Guid.NewGuid(), cancellationToken);
        }
        catch (Exception exception) when (exception is not OperationCanceledException)
        { logger.LogWarning("No se pudo registrar la auditoría del cierre de sesión HYT."); }
        await HttpContext.SignOutAsync("HytCookie");
        return LocalRedirect("/Acceso");
    }
    [AllowAnonymous, HttpGet("/Denegado")]
    public IActionResult Denied() { Response.StatusCode = 403; return View("~/Views/Shared/Denied.cshtml"); }
    private IActionResult LoginView(AccessViewModel model)
    {
        model.Password="";ModelState.Remove(nameof(model.Password));
        model.CaptchaAnswer=null;ModelState.Remove(nameof(model.CaptchaAnswer));
        if(turnstile.Required)
        {
            model.UseManualCaptcha=model.UseManualCaptcha||string.IsNullOrWhiteSpace(turnstile.SiteKey);
            model.CaptchaId=captcha.Create(HttpContext);
        }
        ModelState.Remove(nameof(model.CaptchaId));ModelState.Remove(nameof(model.UseManualCaptcha));
        ViewData["TurnstileRequired"]=turnstile.Required;ViewData["TurnstileSiteKey"]=turnstile.SiteKey;
        return View("Login",model);
    }}
