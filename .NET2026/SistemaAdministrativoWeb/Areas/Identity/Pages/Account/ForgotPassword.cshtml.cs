using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Email;
using SistemaAdministrativoWeb.Infrastructure.Security;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ForgotPasswordModel(
    UserManager<IdentityUser> userManager,
    IAccountEmailService accountEmailService,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<ForgotPasswordModel> logger) : PageModel
{
    // Firma: FRANCO LARA - 31/08/2026 | Implementa recuperacion segura de contrasena y reenvio de confirmacion por Brevo.
    private const string ForgotPasswordAttemptsSessionKey = "Auth:ForgotPasswordAttempts";
    private const string ForgotPasswordCaptchaScope = "FORGOTPWD";

    [BindProperty]
    public InputModel Input { get; set; } = new();

    [BindProperty]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }

    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarTurnstile { get; private set; }
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;
    public bool EmailsEnabled => accountEmailService.IsEnabled;

    public sealed class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
        public string Email { get; set; } = string.Empty;
    }

    public void OnGet()
    {
        PrepareChallenge();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        PrepareChallenge();
        if (!EmailsEnabled)
        {
            ModelState.AddModelError(string.Empty, "La recuperacion por correo no esta disponible temporalmente.");
            return Page();
        }

        if (!ModelState.IsValid)
        {
            return Page();
        }

        var attempts = IncrementAttemptCount();
        if (attempts >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge)
            && !await ValidateChallengeAsync())
        {
            MostrarTurnstile = true;
            return Page();
        }

        var email = Input.Email.Trim();
        var user = await userManager.FindByEmailAsync(email);
        if (user is not null)
        {
            await TrySendAccountEmailAsync(user, email);
        }

        if (attempts >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge))
        {
            HttpContext.Session.Remove(ForgotPasswordAttemptsSessionKey);
        }

        return RedirectToPage("./ForgotPasswordConfirmation");
    }

    private async Task TrySendAccountEmailAsync(IdentityUser user, string email)
    {
        try
        {
            if (!await userManager.IsEmailConfirmedAsync(user))
            {
                var confirmationCode = await userManager.GenerateEmailConfirmationTokenAsync(user);
                var encodedConfirmationCode = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(confirmationCode));
                var confirmationUrl = Url.Page(
                    "/Account/ConfirmEmail",
                    pageHandler: null,
                    values: new { area = "Identity", userId = user.Id, code = encodedConfirmationCode },
                    protocol: Request.Scheme);

                if (!string.IsNullOrWhiteSpace(confirmationUrl))
                {
                    await accountEmailService.SendConfirmationEmailAsync(
                        email,
                        email,
                        confirmationUrl,
                        HttpContext.RequestAborted);
                }

                return;
            }

            var resetCode = await userManager.GeneratePasswordResetTokenAsync(user);
            var encodedResetCode = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(resetCode));
            var resetUrl = Url.Page(
                "/Account/ResetPassword",
                pageHandler: null,
                values: new { area = "Identity", code = encodedResetCode, email },
                protocol: Request.Scheme);

            if (!string.IsNullOrWhiteSpace(resetUrl))
            {
                await accountEmailService.SendResetPasswordEmailAsync(
                    email,
                    email,
                    resetUrl,
                    HttpContext.RequestAborted);
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo procesar el correo de recuperacion para {UserId}.", user.Id);
        }
    }

    private void PrepareChallenge()
    {
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = GetAttemptCount() >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge);
        MostrarCaptchaManual = MostrarTurnstile
            && (UsarCaptchaManual || string.IsNullOrWhiteSpace(TurnstileSiteKey));

        if (MostrarCaptchaManual)
        {
            CaptchaManualCodigo = ManualCaptchaChallengeStore.GetOrCreate(HttpContext, ForgotPasswordCaptchaScope);
        }
        else
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
            CaptchaManualCodigo = string.Empty;
        }
    }

    private async Task<bool> ValidateChallengeAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, ForgotPasswordCaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
                return true;
            }

            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = Request.Form["cf-turnstile-response"].ToString().Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            UsarCaptchaManual = true;
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
            ModelState.AddModelError(string.Empty, "Completa la verificacion automatica o usa el captcha manual.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var result = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (result.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
            return true;
        }

        UsarCaptchaManual = true;
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion automatica. Usa el captcha manual.");
        return false;
    }

    private int GetAttemptCount()
        => HttpContext.Session.GetInt32(ForgotPasswordAttemptsSessionKey).GetValueOrDefault();

    private int IncrementAttemptCount()
    {
        var next = GetAttemptCount() + 1;
        HttpContext.Session.SetInt32(ForgotPasswordAttemptsSessionKey, next);
        return next;
    }
}
