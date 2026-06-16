using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class LoginModel(
    SignInManager<IdentityUser> signInManager,
    UserManager<IdentityUser> userManager,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<LoginModel> logger) : PageModel
{
    private const string LoginFailuresSessionKey = "Auth:LoginFailures";

    [BindProperty]
    public InputModel Input { get; set; } = new();

    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string ReturnUrl { get; set; } = string.Empty;
    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool MostrarTurnstile { get; private set; }

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
        if (User.Identity?.IsAuthenticated == true)
        {
            Response.Redirect(Url.Content("~/"));
            return;
        }

        await HttpContext.SignOutAsync(IdentityConstants.ExternalScheme);
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();
        ReturnUrl = returnUrl ?? Url.Content("~/");
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ReturnUrl = returnUrl;
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        MostrarTurnstile = DebeMostrarTurnstile();

        if (!ModelState.IsValid)
        {
            return Page();
        }

        if (DebeValidarTurnstileEnLogin() && !await ValidarTurnstileAsync())
        {
            IncrementarContador(LoginFailuresSessionKey);
            MostrarTurnstile = true;
            return Page();
        }

        var result = await signInManager.PasswordSignInAsync(
            Input.Email,
            Input.Password,
            Input.RememberMe,
            lockoutOnFailure: true);

        if (result.Succeeded)
        {
            ReiniciarContador(LoginFailuresSessionKey);
            logger.LogInformation("Usuario autenticado.");
            return LocalRedirect(returnUrl);
        }

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
            logger.LogInformation("Usuario inicio sesion con {Provider}.", info.LoginProvider);
            return LocalRedirect(returnUrl);
        }

        var email = info.Principal.FindFirstValue(ClaimTypes.Email);
        if (string.IsNullOrWhiteSpace(email))
        {
            return RedirectToPage("./Login", new { returnUrl });
        }

        email = email.Trim();
        var user = await userManager.FindByEmailAsync(email);
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

            await cuentaAdministradoraRepository.GuardarPerfilUsuarioAsync(new UsuarioPerfilRequest
            {
                AspNetUserId = user.Id,
                NombreCompleto = nombreCompleto,
                CorreoReferencia = email,
                UsuarioRegistro = email
            });
        }

        var addLoginResult = await userManager.AddLoginAsync(user, info);
        if (!addLoginResult.Succeeded && !addLoginResult.Errors.Any(x => string.Equals(x.Code, "LoginAlreadyAssociated", StringComparison.OrdinalIgnoreCase)))
        {
            return RedirectToPage("./Login", new { returnUrl });
        }

        await signInManager.SignInAsync(user, isPersistent: false);
        logger.LogInformation("Usuario creo o vinculo cuenta con {Provider}.", info.LoginProvider);
        return LocalRedirect(returnUrl);
    }

    private async Task<bool> ValidarTurnstileAsync()
    {
        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            ModelState.AddModelError(string.Empty, "Completa la verificacion de seguridad.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            return true;
        }

        logger.LogWarning("Turnstile rechazo login. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion de seguridad. Intenta nuevamente.");
        return false;
    }

    private bool DebeMostrarTurnstile()
        => ObtenerContador(LoginFailuresSessionKey) >= Math.Max(1, turnstileOptions.Value.LoginFailuresBeforeChallenge);

    private bool DebeValidarTurnstileEnLogin()
        => DebeMostrarTurnstile();

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
