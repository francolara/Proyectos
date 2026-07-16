using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ForgotPasswordModel(
    UserManager<ApplicationUser> userManager,
    IAccountEmailService accountEmailService,
    ITurnstileValidationService turnstileValidationService,
    Microsoft.Extensions.Options.IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<ForgotPasswordModel> logger) : PageModel
{
    private const string ForgotPasswordAttemptsSessionKey = "Auth:ForgotPasswordAttempts";
    private const string ForgotPasswordCaptchaScope = "FORGOTPWD";

    [BindProperty]
    public InputModel Input { get; set; } = new();

    [BindProperty]
    [StringLength(20, ErrorMessage = "El codigo captcha no puede exceder 20 caracteres.")]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }
    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarTurnstile { get; private set; }
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;

    public class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingresa un correo valido.")]
        public string Email { get; set; } = string.Empty;
    }

    public void OnGet()
    {
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        if (!ModelState.IsValid) return Page();
        var totalIntentos = IncrementarContador(ForgotPasswordAttemptsSessionKey);
        if (DebeValidarTurnstile(totalIntentos) && !await ValidarTurnstileAsync())
        {
            MostrarTurnstile = true;
            return Page();
        }

        var email = (Input.Email ?? string.Empty).Trim();
        var user = await userManager.FindByEmailAsync(email);

        if (user is not null)
        {
            if (await userManager.IsEmailConfirmedAsync(user))
            {
                var code = await userManager.GeneratePasswordResetTokenAsync(user);
                code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
                var callbackUrl = Url.Page(
                    "/Account/ResetPassword",
                    pageHandler: null,
                    values: new { area = "Identity", code, email },
                    protocol: Request.Scheme);

                if (!string.IsNullOrWhiteSpace(callbackUrl))
                {
                    try
                    {
                        await accountEmailService.SendResetPasswordEmailAsync(email, user.Nombres, callbackUrl);
                    }
                    catch (EmailDeliveryException ex)
                    {
                        logger.LogWarning(ex, "No se pudo enviar correo de recuperacion para {Email}.", email);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error no controlado al enviar correo de recuperacion para {Email}.", email);
                    }
                }
            }
            else
            {
                var code = await userManager.GenerateEmailConfirmationTokenAsync(user);
                code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
                var callbackUrl = Url.Page(
                    "/Account/ConfirmEmail",
                    pageHandler: null,
                    values: new { area = "Identity", userId = user.Id, code },
                    protocol: Request.Scheme);

                if (!string.IsNullOrWhiteSpace(callbackUrl))
                {
                    try
                    {
                        await accountEmailService.SendConfirmationEmailAsync(email, user.Nombres, callbackUrl);
                    }
                    catch (EmailDeliveryException ex)
                    {
                        logger.LogWarning(ex, "No se pudo reenviar confirmacion desde forgot password para {Email}.", email);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error no controlado al reenviar confirmacion desde forgot password para {Email}.", email);
                    }
                }
            }
        }

        return RedirectToPage("./ForgotPasswordConfirmation");
    }

    private async Task<bool> ValidarTurnstileAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, ForgotPasswordCaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
                return true;
            }

            UsarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            UsarCaptchaManual = true;
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
            ModelState.AddModelError(string.Empty, "No se pudo validar Turnstile. Usa el codigo manual de respaldo.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
            return true;
        }

        logger.LogWarning("Turnstile rechazo forgot password. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        UsarCaptchaManual = true;
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, ForgotPasswordCaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar Turnstile. Usa el codigo manual de respaldo.");
        return false;
    }

    private bool DebeMostrarTurnstile()
        => DebeValidarTurnstile(ObtenerContador(ForgotPasswordAttemptsSessionKey));

    private bool DebeValidarTurnstile(int intentos)
        => intentos >= Math.Max(1, turnstileOptions.Value.ForgotPasswordAttemptsBeforeChallenge);

    private void ConfigurarCaptchaManual()
    {
        MostrarCaptchaManual = MostrarTurnstile
            && (UsarCaptchaManual || string.IsNullOrWhiteSpace(TurnstileSiteKey));

        if (MostrarCaptchaManual)
        {
            CaptchaManualCodigo = ManualCaptchaChallengeStore.GetOrCreate(HttpContext, ForgotPasswordCaptchaScope);
            return;
        }

        ManualCaptchaChallengeStore.Clear(HttpContext, ForgotPasswordCaptchaScope);
        CaptchaManualCodigo = string.Empty;
    }

    private int ObtenerContador(string key)
    {
        var valor = HttpContext.Session.GetInt32(key);
        return valor.GetValueOrDefault();
    }

    private int IncrementarContador(string key)
    {
        var siguiente = ObtenerContador(key) + 1;
        HttpContext.Session.SetInt32(key, siguiente);
        return siguiente;
    }

    private void ReiniciarContador(string key)
    {
        HttpContext.Session.Remove(key);
    }
}
