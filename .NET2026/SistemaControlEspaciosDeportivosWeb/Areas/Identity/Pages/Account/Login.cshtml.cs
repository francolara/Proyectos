using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;
using SistemaControlEspaciosDeportivosWeb.Configuration;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class LoginModel(
    SignInManager<ApplicationUser> signInManager,
    UserManager<ApplicationUser> userManager,
    IAccountEmailService accountEmailService,
    ISportCenterStoredProcedureService spService,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    IOptions<IdentityBehaviorSettings> identityBehaviorOptions,
    ILogger<LoginModel> logger) : PageModel
{
    // Firma: FRANCO LARA - 21/07/2026 | Respeta la confirmacion de cuenta configurada y oculta el reenvio cuando la autoconfirmacion esta activa.
    private const string LoginFailuresSessionKey = "Auth:LoginFailures";
    private const string ResendAttemptsSessionKey = "Auth:ResendAttempts";
    private const string LoginCaptchaScope = "LOGIN";

    [BindProperty]
    public InputModel Input { get; set; } = new();

    [BindProperty]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }

    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string ReturnUrl { get; set; } = string.Empty;
    public WebBannerPublicoViewModel? BannerLateral { get; set; }
    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarTurnstile { get; private set; }
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;
    public bool RequiereConfirmacionCorreo =>
        identityBehaviorOptions.Value.RequireConfirmedAccount
        && !identityBehaviorOptions.Value.AutoConfirmEmail;
    public bool CorreosHabilitados => !identityBehaviorOptions.Value.AutoConfirmEmail;

    [TempData]
    public string? ErrorMessage { get; set; }

    [TempData]
    public string? SuccessMessage { get; set; }

    public class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingresa un correo valido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "La contrasena es obligatoria.")]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [Display(Name = "Recordarme")]
        public bool RememberMe { get; set; }
    }

    public async Task OnGetAsync(string? returnUrl = null)
    {
        if (!string.IsNullOrWhiteSpace(ErrorMessage))
        {
            ModelState.AddModelError(string.Empty, ErrorMessage);
        }

        await HttpContext.SignOutAsync(IdentityConstants.ExternalScheme);
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        ReturnUrl = returnUrl ?? Url.Content("~/");
        await CargarBannerLateralAsync();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        if (!ModelState.IsValid)
        {
            await CargarBannerLateralAsync();
            return Page();
        }
        if (DebeValidarTurnstileEnLogin() && !await ValidarTurnstileAsync())
        {
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = true;
            await CargarBannerLateralAsync();
            return Page();
        }

        var email = (Input.Email ?? string.Empty).Trim();
        Input.Email = email;
        var user = await userManager.FindByEmailAsync(email);
        if (user is null)
        {
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(string.Empty, "Intento de inicio de sesion no valido.");
            await CargarBannerLateralAsync();
            return Page();
        }

        if (RequiereConfirmacionCorreo && !await userManager.IsEmailConfirmedAsync(user))
        {
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(string.Empty, "Tu cuenta aun no esta confirmada. Revisa tu correo o reenvia el enlace de confirmacion.");
            await CargarBannerLateralAsync();
            return Page();
        }

        var result = await signInManager.CheckPasswordSignInAsync(user, Input.Password, lockoutOnFailure: true);
        if (result.Succeeded)
        {
            ReiniciarContador(LoginFailuresSessionKey);
            var propiedadesAutenticacion = new AuthenticationProperties
            {
                IsPersistent = Input.RememberMe,
                AllowRefresh = true,
                ExpiresUtc = DateTimeOffset.UtcNow.Add(Input.RememberMe ? TimeSpan.FromDays(2) : TimeSpan.FromMinutes(30))
            };

            await signInManager.SignInAsync(user, propiedadesAutenticacion);
            logger.LogInformation("Usuario inicio sesion.");
            if (!string.IsNullOrWhiteSpace(returnUrl)
                && Url.IsLocalUrl(returnUrl)
                && !string.Equals(returnUrl, Url.Content("~/"), StringComparison.OrdinalIgnoreCase))
            {
                return LocalRedirect(returnUrl);
            }

            var negocios = await spService.PanelListarNegociosUsuarioAsync(user.Id);
            if (negocios.Count > 0)
            {
                return RedirectToAction("Index", "Panel", new { negocioId = negocios[0].NegocioId });
            }

            return RedirectToAction("Index", "Home");
        }

        if (result.RequiresTwoFactor)
        {
            return RedirectToPage("./LoginWith2fa", new { ReturnUrl = returnUrl, Input.RememberMe });
        }

        if (result.IsLockedOut)
        {
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = true;
            logger.LogWarning("Cuenta bloqueada.");
            return RedirectToPage("./Lockout");
        }

        IncrementarContador(LoginFailuresSessionKey);
        MostrarTurnstile = DebeMostrarTurnstile();
        ModelState.AddModelError(string.Empty, "Intento de inicio de sesion no valido.");
        await CargarBannerLateralAsync();
        return Page();
    }

    public async Task<IActionResult> OnPostResendConfirmationAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        ReturnUrl = returnUrl;

        if (!RequiereConfirmacionCorreo)
        {
            SuccessMessage = "La confirmacion de correo no es necesaria con la configuracion actual.";
            await CargarBannerLateralAsync();
            return Page();
        }

        var email = (Input.Email ?? string.Empty).Trim();
        Input.Email = email;
        if (string.IsNullOrWhiteSpace(email))
        {
            ModelState.AddModelError(string.Empty, "Ingresa tu correo para reenviar la confirmacion.");
            await CargarBannerLateralAsync();
            return Page();
        }
        var totalReenvios = IncrementarContador(ResendAttemptsSessionKey);
        if (DebeValidarTurnstileEnReenvio(totalReenvios) && !await ValidarTurnstileAsync())
        {
            MostrarTurnstile = true;
            await CargarBannerLateralAsync();
            return Page();
        }

        var correoEnviado = false;
        var user = await userManager.FindByEmailAsync(email);
        if (user is not null && !await userManager.IsEmailConfirmedAsync(user))
        {
            var code = await userManager.GenerateEmailConfirmationTokenAsync(user);
            code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
            var callbackUrl = Url.Page(
                "/Account/ConfirmEmail",
                pageHandler: null,
                values: new { area = "Identity", userId = user.Id, code, returnUrl },
                protocol: Request.Scheme);

            if (!string.IsNullOrWhiteSpace(callbackUrl))
            {
                try
                {
                    await accountEmailService.SendConfirmationEmailAsync(user.Email ?? email, user.Nombres, callbackUrl);
                    correoEnviado = true;
                }
                catch (EmailDeliveryException ex)
                {
                    logger.LogWarning(ex, "No se pudo reenviar correo de confirmacion para {Email}.", email);
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "Error no controlado al reenviar correo de confirmacion para {Email}.", email);
                }
            }
        }
        else
        {
            correoEnviado = true;
        }

        if (correoEnviado)
        {
            SuccessMessage = "Si tu cuenta existe y aun no esta confirmada, te enviamos un nuevo correo de confirmacion.";
        }
        else
        {
            ModelState.AddModelError(string.Empty, "No se pudo enviar el correo de confirmacion en este momento. Intenta nuevamente en unos minutos.");
        }

        await CargarBannerLateralAsync();
        MostrarTurnstile = DebeMostrarTurnstile();
        return Page();
    }

    public IActionResult OnPostExternalLogin(string provider, string? returnUrl = null, string? flow = null)
    {
        returnUrl ??= Url.Content("~/");
        flow = string.Equals(flow, "register", StringComparison.OrdinalIgnoreCase) ? "register" : "login";
        var redirectUrl = Url.Page("./Login", pageHandler: "ExternalLoginCallback", values: new { returnUrl, flow });
        var properties = signInManager.ConfigureExternalAuthenticationProperties(provider, redirectUrl);
        return Challenge(properties, provider);
    }

    public async Task<IActionResult> OnGetExternalLoginCallbackAsync(string? returnUrl = null, string? remoteError = null, string? flow = null)
    {
        returnUrl ??= Url.Content("~/");
        flow = string.Equals(flow, "register", StringComparison.OrdinalIgnoreCase) ? "register" : "login";

        if (!string.IsNullOrWhiteSpace(remoteError))
        {
            ErrorMessage = $"Error del proveedor externo: {remoteError}";
            return RedirectToPage("./Login", new { returnUrl });
        }

        var info = await signInManager.GetExternalLoginInfoAsync();
        if (info is null)
        {
            ErrorMessage = "No se pudo cargar la informacion del proveedor externo.";
            return RedirectToPage("./Login", new { returnUrl });
        }

        var result = await signInManager.ExternalLoginSignInAsync(info.LoginProvider, info.ProviderKey, isPersistent: false, bypassTwoFactor: true);
        if (result.Succeeded)
        {
            logger.LogInformation("Usuario inicio sesion con {Provider}.", info.LoginProvider);
            return await RedirigirPostLoginAsync(returnUrl, flow);
        }

        if (result.IsLockedOut)
        {
            logger.LogWarning("Cuenta bloqueada.");
            return RedirectToPage("./Lockout");
        }

        var email = info.Principal.FindFirstValue(ClaimTypes.Email);
        if (string.IsNullOrWhiteSpace(email))
        {
            ErrorMessage = "El proveedor externo no devolvio un correo electronico valido.";
            return RedirectToPage("./Login", new { returnUrl });
        }

        email = email.Trim();
        var user = await userManager.FindByEmailAsync(email);
        if (user is null)
        {
            var nombreCompleto = (info.Principal.FindFirstValue(ClaimTypes.Name) ?? string.Empty).Trim();
            var givenName = (info.Principal.FindFirstValue(ClaimTypes.GivenName) ?? string.Empty).Trim();
            var surname = (info.Principal.FindFirstValue(ClaimTypes.Surname) ?? string.Empty).Trim();
            var nombres = !string.IsNullOrWhiteSpace(givenName) ? givenName : nombreCompleto;
            var apellidos = !string.IsNullOrWhiteSpace(surname) ? surname : (string.IsNullOrWhiteSpace(nombreCompleto) ? "Usuario" : nombreCompleto);

            user = new ApplicationUser
            {
                UserName = email,
                Email = email,
                EmailConfirmed = identityBehaviorOptions.Value.AutoConfirmEmail,
                Nombres = string.IsNullOrWhiteSpace(nombres) ? email : nombres
            };

            var createResult = await userManager.CreateAsync(user);
            if (!createResult.Succeeded)
            {
                ErrorMessage = "No se pudo crear la cuenta con el proveedor externo.";
                return RedirectToPage("./Login", new { returnUrl });
            }

            try
            {
                await spService.UsuariosPublicosGuardarPerfilAsync(new UsuarioPublicoPerfilViewModel
                {
                    UsuarioId = user.Id,
                    TipoDocumento = "0",
                    Nombres = user.Nombres ?? string.Empty,
                    Apellidos = apellidos,
                    Correo = email
                }, email);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "No se pudo sincronizar perfil publico para login externo de {Email}.", email);
            }
        }

        var addLoginResult = await userManager.AddLoginAsync(user, info);
        if (!addLoginResult.Succeeded && !addLoginResult.Errors.Any(x => string.Equals(x.Code, "LoginAlreadyAssociated", StringComparison.OrdinalIgnoreCase)))
        {
            ErrorMessage = "No se pudo vincular la cuenta externa.";
            return RedirectToPage("./Login", new { returnUrl });
        }

        await signInManager.SignInAsync(user, isPersistent: false);
        logger.LogInformation("Usuario creo o vinculo cuenta con {Provider}.", info.LoginProvider);
        return await RedirigirPostLoginAsync(returnUrl, flow);
    }

    private async Task CargarBannerLateralAsync()
    {
        try
        {
            BannerLateral = await spService.WebBannersObtenerFijoPorTipoAsync((int)BannerTipo.Login);
        }
        catch
        {
            BannerLateral = null;
        }
    }

    private async Task<IActionResult> RedirigirPostLoginAsync(string? returnUrl, string? flow)
    {
        if (!string.IsNullOrWhiteSpace(returnUrl)
            && Url.IsLocalUrl(returnUrl)
            && !string.Equals(returnUrl, Url.Content("~/"), StringComparison.OrdinalIgnoreCase))
        {
            return LocalRedirect(returnUrl);
        }

        if (string.Equals(flow, "register", StringComparison.OrdinalIgnoreCase))
        {
            return RedirectToAction("Index", "Home");
        }

        var usuario = await userManager.GetUserAsync(User);
        if (usuario is not null)
        {
            var negocios = await spService.PanelListarNegociosUsuarioAsync(usuario.Id);
            if (negocios.Count > 0)
            {
                return RedirectToAction("Index", "Panel", new { negocioId = negocios[0].NegocioId });
            }
        }

        return RedirectToAction("Index", "Home");
    }

    private async Task<bool> ValidarTurnstileAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, LoginCaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, LoginCaptchaScope);
                return true;
            }

            UsarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            UsarCaptchaManual = true;
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
            ModelState.AddModelError(string.Empty, "No se pudo validar Turnstile. Usa el codigo manual de respaldo.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, LoginCaptchaScope);
            return true;
        }

        logger.LogWarning("Turnstile rechazo login/reenvio confirmacion. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        UsarCaptchaManual = true;
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar Turnstile. Usa el codigo manual de respaldo.");
        return false;
    }

    private void ConfigurarCaptchaManual()
    {
        MostrarCaptchaManual = MostrarTurnstile
            && (UsarCaptchaManual || string.IsNullOrWhiteSpace(TurnstileSiteKey));

        if (MostrarCaptchaManual)
        {
            CaptchaManualCodigo = ManualCaptchaChallengeStore.GetOrCreate(HttpContext, LoginCaptchaScope);
            return;
        }

        ManualCaptchaChallengeStore.Clear(HttpContext, LoginCaptchaScope);
        CaptchaManualCodigo = string.Empty;
    }

    private bool DebeMostrarTurnstile()
    {
        var loginFailures = ObtenerContador(LoginFailuresSessionKey);
        var resendAttempts = ObtenerContador(ResendAttemptsSessionKey);
        return DebeValidarTurnstileEnLogin(loginFailures) || DebeValidarTurnstileEnReenvio(resendAttempts);
    }

    private bool DebeValidarTurnstileEnLogin()
        => DebeValidarTurnstileEnLogin(ObtenerContador(LoginFailuresSessionKey));

    private bool DebeValidarTurnstileEnLogin(int loginFailures)
        => loginFailures >= Math.Max(1, turnstileOptions.Value.LoginFailuresBeforeChallenge);

    private bool DebeValidarTurnstileEnReenvio(int resendAttempts)
        => resendAttempts >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge);

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
