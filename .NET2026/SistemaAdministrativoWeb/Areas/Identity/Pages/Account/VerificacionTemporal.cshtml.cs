using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Security;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[Authorize]
public class VerificacionTemporalModel(
    UserManager<IdentityUser> userManager,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<VerificacionTemporalModel> logger) : PageModel
{
    private const string CaptchaScope = "TEMPVERIFY";

    [BindProperty]
    [StringLength(20, ErrorMessage = "El codigo captcha no puede exceder 20 caracteres.")]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }

    [BindProperty(SupportsGet = true)]
    public string ReturnUrl { get; set; } = string.Empty;

    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;

    public async Task<IActionResult> OnGetAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        var user = await userManager.GetUserAsync(User);
        if (user is null || !await RequiereCambioContrasenaTemporalAsync(user))
        {
            return LocalRedirect(Url.Content("~/"));
        }

        if (UsuarioYaVerificado(user))
        {
            return RedirectToPage("./CambiarContrasenaTemporal", new { returnUrl = ReturnUrl });
        }

        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        ConfigurarCaptchaManual();
        return Page();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        var user = await userManager.GetUserAsync(User);
        if (user is null || !await RequiereCambioContrasenaTemporalAsync(user))
        {
            return LocalRedirect(Url.Content("~/"));
        }

        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        ConfigurarCaptchaManual();

        if (!ModelState.IsValid)
        {
            return Page();
        }

        if (!await ValidarDesafioAccesoAsync())
        {
            return Page();
        }

        HttpContext.Session.SetString(TemporaryPasswordFlowConstants.VerificationSessionKey, user.Id);
        return RedirectToPage("./CambiarContrasenaTemporal", new { returnUrl = ReturnUrl });
    }

    public async Task<IActionResult> OnPostFallbackAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        var user = await userManager.GetUserAsync(User);
        if (user is null || !await RequiereCambioContrasenaTemporalAsync(user))
        {
            return LocalRedirect(Url.Content("~/"));
        }

        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        UsarCaptchaManual = true;
        ModelState.Clear();
        ConfigurarCaptchaManual();
        return Page();
    }

    private async Task<bool> ValidarDesafioAccesoAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, CaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, CaptchaScope);
                return true;
            }

            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, CaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, CaptchaScope);
            ModelState.AddModelError(string.Empty, "Completa la verificacion automatica o usa el captcha manual de respaldo.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, CaptchaScope);
            return true;
        }

        logger.LogWarning("Turnstile rechazo verificacion temporal. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, CaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion automatica. Usa el captcha manual de respaldo.");
        return false;
    }

    private void ConfigurarCaptchaManual()
    {
        MostrarCaptchaManual = UsarCaptchaManual || string.IsNullOrWhiteSpace(TurnstileSiteKey);

        if (MostrarCaptchaManual)
        {
            CaptchaManualCodigo = ManualCaptchaChallengeStore.GetOrCreate(HttpContext, CaptchaScope);
            return;
        }

        ManualCaptchaChallengeStore.Clear(HttpContext, CaptchaScope);
        CaptchaManualCodigo = string.Empty;
    }

    private bool UsuarioYaVerificado(IdentityUser user)
        => string.Equals(
            HttpContext.Session.GetString(TemporaryPasswordFlowConstants.VerificationSessionKey),
            user.Id,
            StringComparison.Ordinal);

    private async Task<bool> RequiereCambioContrasenaTemporalAsync(IdentityUser user)
    {
        var claims = await userManager.GetClaimsAsync(user);
        return claims.Any(x =>
            string.Equals(x.Type, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType, StringComparison.Ordinal)
            && string.Equals(x.Value, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimValue, StringComparison.Ordinal));
    }
}
