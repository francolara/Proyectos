using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Email;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class LoginModel(
    SignInManager<IdentityUser> signInManager,
    UserManager<IdentityUser> userManager,
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    IAccountEmailService accountEmailService,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    IOptions<IdentityBehaviorSettings> identityBehaviorOptions,
    ILogger<LoginModel> logger) : PageModel
{
    // Firma: FRANCO LARA - 31/08/2026 | Agrega recuperacion y reenvio mediante Brevo y diferencia visualmente los avisos de entrega fallida.
    // Firma: FRANCO LARA - 02/09/2026 | Aplica 30 minutos a la sesion normal y 2 dias a Recordarme sin perder bloqueo, segundo factor, captcha ni redireccion por empresa.
    private const string LoginFailuresSessionKey = "Auth:LoginFailures";
    private const string ResendAttemptsSessionKey = "Auth:ResendConfirmationAttempts";
    private const string LoginCaptchaScope = "LOGIN";
    private static readonly string LoginDebugPath = Path.Combine(AppContext.BaseDirectory, "login-debug.log");

    [BindProperty]
    public InputModel Input { get; set; } = new();

    [BindProperty]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }

    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string ReturnUrl { get; set; } = string.Empty;
    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarTurnstile { get; private set; }
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;
    public bool EmailsEnabled => accountEmailService.IsEnabled;
    public bool RequiresEmailConfirmation =>
        identityBehaviorOptions.Value.RequireConfirmedAccount
        && !identityBehaviorOptions.Value.AutoConfirmEmail;

    [TempData]
    public string? AccountSuccessMessage { get; set; }

    [TempData]
    public string? AccountMessageType { get; set; }

    public sealed class InputModel
    {
        [Required(ErrorMessage = "Ingrese su correo.")]
        [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese su contrasena.")]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        public bool RememberMe { get; set; }
    }

    public async Task OnGetAsync(string? returnUrl = null)
    {
        RegistrarDebug($"GET login | Auth={User.Identity?.IsAuthenticated} | ReturnUrl={returnUrl}");
        if (User.Identity?.IsAuthenticated == true)
        {
            Response.Redirect(Url.Content("~/"));
            return;
        }

        await HttpContext.SignOutAsync(IdentityConstants.ExternalScheme);
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        ReturnUrl = returnUrl ?? Url.Content("~/");
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        RegistrarDebug($"POST login START | Email={Input.Email} | ReturnUrl={returnUrl} | Auth={User.Identity?.IsAuthenticated}");
        returnUrl ??= Url.Content("~/");
        ReturnUrl = returnUrl;
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();

        if (!ModelState.IsValid)
        {
            RegistrarDebug($"POST login | ModelState invalido | Detalle={ObtenerErroresModelState(ModelState)}");
            return Page();
        }

        if (DebeValidarTurnstileEnLogin() && !await ValidarDesafioAccesoAsync())
        {
            RegistrarDebug("POST login | desafio no valido");
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = true;
            return Page();
        }

        var authenticatedUser = await userManager.FindByEmailAsync(Input.Email.Trim());
        if (authenticatedUser is null)
        {
            RegistrarDebug("POST login | usuario no encontrado por email");
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(string.Empty, "Credenciales invalidas.");
            return Page();
        }

        var result = await signInManager.PasswordSignInAsync(
            authenticatedUser,
            Input.Password,
            Input.RememberMe,
            lockoutOnFailure: true);

        if (result.Succeeded)
        {
            var authenticationProperties = new AuthenticationProperties
            {
                IsPersistent = Input.RememberMe,
                AllowRefresh = true,
                ExpiresUtc = DateTimeOffset.UtcNow.Add(
                    Input.RememberMe
                        ? TimeSpan.FromDays(2)
                        : TimeSpan.FromMinutes(30))
            };
            await signInManager.SignInAsync(authenticatedUser, authenticationProperties);

            RegistrarDebug($"POST login | SUCCESS | UserId={authenticatedUser.Id}");
            ReiniciarContador(LoginFailuresSessionKey);
            currentCompanyAccessor.LimpiarEmpresa();
            logger.LogInformation("Usuario autenticado.");

            if (authenticatedUser is not null
                && await RequiereCambioContrasenaTemporalAsync(authenticatedUser))
            {
                RegistrarDebug($"POST login | requiere cambio temporal | UserId={authenticatedUser.Id}");
                HttpContext.Session.Remove(TemporaryPasswordFlowConstants.VerificationSessionKey);
                return RedirectToPage("./VerificacionTemporal", new { returnUrl });
            }

            RegistrarDebug($"POST login | redireccion contexto | UserId={authenticatedUser!.Id}");
            return await RedirigirSegunContextoAsync(authenticatedUser, returnUrl);
        }

        if (result.IsLockedOut)
        {
            RegistrarDebug($"POST login | LOCKOUT | UserId={authenticatedUser.Id}");
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(string.Empty, "La cuenta se encuentra bloqueada temporalmente por varios intentos fallidos.");
            return Page();
        }

        if (result.IsNotAllowed)
        {
            RegistrarDebug($"POST login | NOT_ALLOWED | UserId={authenticatedUser.Id}");
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(
                string.Empty,
                RequiresEmailConfirmation
                    ? "Debes confirmar tu correo antes de iniciar sesion. Puedes solicitar un nuevo enlace."
                    : "La cuenta no tiene permitido iniciar sesion con la configuracion actual.");
            return Page();
        }

        if (result.RequiresTwoFactor)
        {
            RegistrarDebug($"POST login | REQUIRES_2FA | UserId={authenticatedUser.Id}");
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = DebeMostrarTurnstile();
            ModelState.AddModelError(string.Empty, "La cuenta requiere un segundo factor de autenticacion.");
            return Page();
        }

        RegistrarDebug($"POST login | FAILED | UserId={authenticatedUser.Id}");
        IncrementarContador(LoginFailuresSessionKey);
        MostrarTurnstile = DebeMostrarTurnstile();
        ModelState.AddModelError(string.Empty, "Credenciales invalidas.");
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

    public async Task<IActionResult> OnPostResendConfirmationAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ReturnUrl = returnUrl;
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ConfigurarCaptchaManual();
        ModelState.Remove("Input.Password");

        if (!RequiresEmailConfirmation || !EmailsEnabled)
        {
            ModelState.AddModelError(string.Empty, "El reenvio de confirmacion no esta disponible con la configuracion actual.");
            return Page();
        }

        if (string.IsNullOrWhiteSpace(Input.Email) || !new EmailAddressAttribute().IsValid(Input.Email))
        {
            ModelState.AddModelError("Input.Email", "Ingrese un correo valido.");
            return Page();
        }

        var resendAttempts = IncrementarContador(ResendAttemptsSessionKey);
        if (resendAttempts >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge)
            && !await ValidarDesafioAccesoAsync())
        {
            MostrarTurnstile = true;
            return Page();
        }

        var email = Input.Email.Trim();
        var user = await userManager.FindByEmailAsync(email);
        if (user is not null && !await userManager.IsEmailConfirmedAsync(user))
        {
            var code = await userManager.GenerateEmailConfirmationTokenAsync(user);
            var encodedCode = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
            var callbackUrl = Url.Page(
                "/Account/ConfirmEmail",
                pageHandler: null,
                values: new { area = "Identity", userId = user.Id, code = encodedCode, returnUrl },
                protocol: Request.Scheme);

            if (!string.IsNullOrWhiteSpace(callbackUrl))
            {
                try
                {
                    await accountEmailService.SendConfirmationEmailAsync(
                        email,
                        email,
                        callbackUrl,
                        HttpContext.RequestAborted);
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "No se pudo reenviar confirmacion para {UserId}.", user.Id);
                }
            }
        }

        if (resendAttempts >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge))
        {
            ReiniciarContador(ResendAttemptsSessionKey);
        }

        AccountSuccessMessage = "Si la cuenta existe y esta pendiente, enviamos un nuevo enlace de confirmacion.";
        return RedirectToPage("./Login", new { returnUrl });
    }

    public async Task<IActionResult> OnGetExternalLoginCallbackAsync(string? returnUrl = null, string? remoteError = null, string? flow = null)
    {
        returnUrl ??= Url.Content("~/");
        flow = string.Equals(flow, "register", StringComparison.OrdinalIgnoreCase) ? "register" : "login";

        if (!string.IsNullOrWhiteSpace(remoteError))
        {
            ModelState.AddModelError(string.Empty, $"Error del proveedor externo: {remoteError}");
            return RedirectToPage("./Login", new { returnUrl });
        }

        var info = await signInManager.GetExternalLoginInfoAsync();
        if (info is null)
        {
            ModelState.AddModelError(string.Empty, "No se pudo cargar la informacion del proveedor externo.");
            return RedirectToPage("./Login", new { returnUrl });
        }

        var result = await signInManager.ExternalLoginSignInAsync(info.LoginProvider, info.ProviderKey, isPersistent: false, bypassTwoFactor: true);
        if (result.Succeeded)
        {
            currentCompanyAccessor.LimpiarEmpresa();
            logger.LogInformation("Usuario inicio sesion con {Provider}.", info.LoginProvider);
            var authenticatedUser = await userManager.FindByLoginAsync(info.LoginProvider, info.ProviderKey);
            return await RedirigirSegunContextoAsync(authenticatedUser, returnUrl);
        }

        var email = info.Principal.FindFirstValue(ClaimTypes.Email);
        if (string.IsNullOrWhiteSpace(email))
        {
            return RedirectToPage("./Login", new { returnUrl });
        }

        email = email.Trim();
        var user = await userManager.FindByEmailAsync(email);
        var userCreated = false;
        if (user is null)
        {
            var nombreCompleto = (info.Principal.FindFirstValue(ClaimTypes.Name) ?? email).Trim();

            user = new IdentityUser
            {
                UserName = email,
                Email = email,
                EmailConfirmed = true
            };

            var createResult = await userManager.CreateAsync(user);
            if (!createResult.Succeeded)
            {
                return RedirectToPage("./Login", new { returnUrl });
            }

            userCreated = true;

            await cuentaAdministradoraRepository.GuardarPerfilUsuarioAsync(new UsuarioPerfilRequest
            {
                AspNetUserId = user.Id,
                NombreCompleto = nombreCompleto,
                CorreoReferencia = email,
                UsuarioRegistro = email
            });
        }

        else if (!await userManager.IsEmailConfirmedAsync(user))
        {
            user.EmailConfirmed = true;
            var confirmExternalEmailResult = await userManager.UpdateAsync(user);
            if (!confirmExternalEmailResult.Succeeded)
            {
                return RedirectToPage("./Login", new { returnUrl });
            }

            userCreated = true;
        }

        var addLoginResult = await userManager.AddLoginAsync(user, info);
        if (!addLoginResult.Succeeded && !addLoginResult.Errors.Any(x => string.Equals(x.Code, "LoginAlreadyAssociated", StringComparison.OrdinalIgnoreCase)))
        {
            return RedirectToPage("./Login", new { returnUrl });
        }

        currentCompanyAccessor.LimpiarEmpresa();
        await signInManager.SignInAsync(user, isPersistent: false);
        logger.LogInformation("Usuario creo o vinculo cuenta con {Provider}.", info.LoginProvider);

        if (userCreated && accountEmailService.IsEnabled)
        {
            var loginUrl = Url.Page(
                "/Account/Login",
                pageHandler: null,
                values: new { area = "Identity" },
                protocol: Request.Scheme);
            if (!string.IsNullOrWhiteSpace(loginUrl))
            {
                try
                {
                    await accountEmailService.SendWelcomeEmailAsync(
                        email,
                        info.Principal.FindFirstValue(ClaimTypes.Name),
                        loginUrl,
                        HttpContext.RequestAborted);
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "No se pudo enviar bienvenida por acceso externo para {UserId}.", user.Id);
                }
            }
        }

        return await RedirigirSegunContextoAsync(user, returnUrl);
    }

    private async Task<bool> ValidarDesafioAccesoAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, LoginCaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, LoginCaptchaScope);
                return true;
            }

            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
            ModelState.AddModelError(string.Empty, "Completa la verificacion automatica o usa el captcha manual de respaldo.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, LoginCaptchaScope);
            return true;
        }

        logger.LogWarning("Turnstile rechazo login. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, LoginCaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion automatica. Usa el captcha manual de respaldo.");
        return false;
    }

    private bool DebeMostrarTurnstile()
        => ObtenerContador(LoginFailuresSessionKey) >= Math.Max(1, turnstileOptions.Value.LoginFailuresBeforeChallenge)
            || ObtenerContador(ResendAttemptsSessionKey) >= Math.Max(1, turnstileOptions.Value.ResendAttemptsBeforeChallenge);

    private bool DebeValidarTurnstileEnLogin()
        => DebeMostrarTurnstile();

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

    private async Task<IActionResult> RedirigirSegunContextoAsync(IdentityUser? authenticatedUser, string returnUrl)
    {
        var homeUrl = Url.Content("~/");
        if (Url.IsLocalUrl(returnUrl)
            && !string.Equals(returnUrl, homeUrl, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(returnUrl, "/", StringComparison.OrdinalIgnoreCase))
        {
            return LocalRedirect(returnUrl);
        }

        if (authenticatedUser is null)
        {
            return LocalRedirect(homeUrl);
        }

        if (await userManager.IsInRoleAsync(authenticatedUser, "SuperAdmin"))
        {
            return RedirectToAction("Index", "Plataforma", new { area = "" });
        }

        var contextoLogin = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(authenticatedUser.Id, HttpContext.RequestAborted);
        if (contextoLogin is null || !contextoLogin.TieneAcceso)
        {
            return RedirectToAction("Index", "EmpresaContexto", new { area = "" });
        }

        if (contextoLogin.CantidadEmpresasAsignadas == 1 && contextoLogin.IdEmpresaPredeterminada.HasValue)
        {
            currentCompanyAccessor.EstablecerEmpresa(
                contextoLogin.IdEmpresaPredeterminada.Value,
                contextoLogin.RazonSocialEmpresaPredeterminada ?? "Empresa");
            return RedirectToAction("Index", "Panel", new { area = "" });
        }

        if (contextoLogin.SoloModulosCuenta)
        {
            return RedirectToAction("Index", "Configuracion", new { area = "" });
        }

        return RedirectToAction("Index", "EmpresaContexto", new { area = "" });
    }

    private async Task<bool> RequiereCambioContrasenaTemporalAsync(IdentityUser user)
    {
        var claims = await userManager.GetClaimsAsync(user);
        return claims.Any(x =>
            string.Equals(x.Type, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType, StringComparison.Ordinal)
            && string.Equals(x.Value, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimValue, StringComparison.Ordinal));
    }

    private static void RegistrarDebug(string mensaje)
    {
        try
        {
            System.IO.File.AppendAllText(LoginDebugPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {mensaje}{Environment.NewLine}");
        }
        catch
        {
        }
    }

    private static string ObtenerErroresModelState(ModelStateDictionary modelState)
    {
        var errores = modelState
            .Where(x => x.Value is not null && x.Value.Errors.Count > 0)
            .Select(x =>
            {
                var mensajes = x.Value!.Errors
                    .Select(error => string.IsNullOrWhiteSpace(error.ErrorMessage)
                        ? error.Exception?.Message ?? "(sin mensaje)"
                        : error.ErrorMessage)
                    .ToArray();

                return $"{x.Key}=[{string.Join(" | ", mensajes)}]";
            })
            .ToArray();

        return errores.Length == 0 ? "(sin errores detallados)" : string.Join("; ", errores);
    }
}
