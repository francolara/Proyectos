using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class RegisterModel(
    UserManager<IdentityUser> userManager,
    SignInManager<IdentityUser> signInManager,
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    IOptions<IdentityBehaviorSettings> identityBehaviorOptions,
    ILogger<RegisterModel> logger) : PageModel
{
    [BindProperty]
    public UsuarioRegistroInput Usuario { get; set; } = new();

    [BindProperty]
    public string? CaptchaManual { get; set; }

    [BindProperty]
    public bool UsarCaptchaManual { get; set; }

    [BindProperty(SupportsGet = true)]
    public string Plan { get; set; } = string.Empty;

    public string ReturnUrl { get; set; } = string.Empty;
    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string TurnstileSiteKey { get; private set; } = string.Empty;
    public bool LimpiarCamposIniciales { get; private set; }
    public bool MostrarCaptchaManual { get; private set; }
    public string CaptchaManualCodigo { get; private set; } = string.Empty;
    private const string RegisterCaptchaScope = "REGISTER";

    public sealed class UsuarioRegistroInput
    {
        [Required(ErrorMessage = "Ingrese su nombre completo.")]
        [StringLength(180, ErrorMessage = "El nombre completo no puede exceder 180 caracteres.")]
        public string NombreCompleto { get; set; } = string.Empty;

        [StringLength(30, ErrorMessage = "El telefono no puede exceder 30 caracteres.")]
        public string? Telefono { get; set; }

        [Required(ErrorMessage = "Ingrese su correo.")]
        [EmailAddress(ErrorMessage = "Ingrese un correo electronico valido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese su contrasena.")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "La contrasena debe tener entre 6 y 100 caracteres.")]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [Required(ErrorMessage = "Confirme su contrasena.")]
        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "Las contrasenas no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;
    }

    public async Task OnGetAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        Plan = NormalizarPlan(Plan);
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        ConfigurarCaptchaManual();
        LimpiarCamposIniciales = true;
    }

    public async Task<IActionResult> OnPostUsuarioAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        Plan = NormalizarPlan(Plan);
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        ConfigurarCaptchaManual();

        if (!TryValidateModel(Usuario, nameof(Usuario)))
        {
            return Page();
        }

        if (!await ValidarDesafioAccesoAsync())
        {
            return Page();
        }

        var email = Usuario.Email.Trim();
        var user = new IdentityUser
        {
            UserName = email,
            Email = email,
            EmailConfirmed = identityBehaviorOptions.Value.AutoConfirmEmail
        };

        var result = await userManager.CreateAsync(user, Usuario.Password);
        if (!result.Succeeded)
        {
            AgregarErrores(result);
            return Page();
        }

        await cuentaAdministradoraRepository.GuardarPerfilUsuarioAsync(new UsuarioPerfilRequest
        {
            AspNetUserId = user.Id,
            NombreCompleto = Usuario.NombreCompleto.Trim(),
            Telefono = LimpiarTelefono(Usuario.Telefono),
            CorreoReferencia = user.Email,
            UsuarioRegistro = user.Email
        });

        logger.LogInformation("Usuario simple registrado.");
        currentCompanyAccessor.LimpiarEmpresa();
        await signInManager.SignInAsync(user, isPersistent: false);
        return LocalRedirect(ReturnUrl);
    }

    public IActionResult OnPostExternalLogin(string provider, string? returnUrl = null, string? flow = null)
    {
        returnUrl ??= Url.Content("~/");
        flow = string.Equals(flow, "register", StringComparison.OrdinalIgnoreCase) ? "register" : "login";
        var redirectUrl = Url.Page("./Login", pageHandler: "ExternalLoginCallback", values: new { returnUrl, flow });
        var properties = signInManager.ConfigureExternalAuthenticationProperties(provider, redirectUrl);
        return Challenge(properties, provider);
    }

    private void AgregarErrores(IdentityResult result)
    {
        foreach (var error in result.Errors)
        {
            ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error));
        }
    }

    private static string TraducirErrorIdentity(IdentityError error)
    {
        return IdentityErrorTranslator.Translate(error);
    }

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
    }

    public string ObtenerNombrePlanSeleccionado()
    {
        return Plan switch
        {
            "GRATIS" => "Gratis",
            "EMPRENDEDOR" => "Emprendedor",
            "CONTADOR" => "Contador",
            _ => "Registro general"
        };
    }

    public string ObtenerResumenPlanSeleccionado()
    {
        return Plan switch
        {
            "GRATIS" => "Prueba inicial de 30 dias con 1 empresa y 1 usuario.",
            "EMPRENDEDOR" => "Plan mensual para pequenas empresas con hasta 3 empresas y 3 usuarios.",
            "CONTADOR" => "Plan mensual recomendado para estudios contables y gestion multiempresa.",
            _ => "Crea tu cuenta y luego vincula el plan comercial que corresponda."
        };
    }

    private static string NormalizarPlan(string? plan)
    {
        var valor = (plan ?? string.Empty).Trim().ToUpperInvariant();
        return valor switch
        {
            "GRATIS" => valor,
            "EMPRENDEDOR" => valor,
            "CONTADOR" => valor,
            _ => string.Empty
        };
    }

    private async Task<bool> ValidarDesafioAccesoAsync()
    {
        if (MostrarCaptchaManual)
        {
            if (ManualCaptchaChallengeStore.Validate(HttpContext, RegisterCaptchaScope, CaptchaManual))
            {
                ManualCaptchaChallengeStore.Clear(HttpContext, RegisterCaptchaScope);
                return true;
            }

            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, RegisterCaptchaScope);
            ModelState.AddModelError(string.Empty, "El codigo captcha manual no es valido.");
            return false;
        }

        var token = (Request.Form["cf-turnstile-response"].ToString() ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            MostrarCaptchaManual = true;
            CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, RegisterCaptchaScope);
            ModelState.AddModelError(string.Empty, "Completa la verificacion automatica o usa el captcha manual de respaldo.");
            return false;
        }

        var remoteIp = HttpContext.Connection.RemoteIpAddress?.ToString();
        var resultado = await turnstileValidationService.VerifyAsync(token, remoteIp, HttpContext.RequestAborted);
        if (resultado.Success)
        {
            ManualCaptchaChallengeStore.Clear(HttpContext, RegisterCaptchaScope);
            return true;
        }

        logger.LogWarning("Turnstile rechazo registro. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        MostrarCaptchaManual = true;
        CaptchaManualCodigo = ManualCaptchaChallengeStore.Refresh(HttpContext, RegisterCaptchaScope);
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion automatica. Usa el captcha manual de respaldo.");
        return false;
    }

    private void ConfigurarCaptchaManual()
    {
        MostrarCaptchaManual = UsarCaptchaManual || string.IsNullOrWhiteSpace(TurnstileSiteKey);

        if (MostrarCaptchaManual)
        {
            CaptchaManualCodigo = ManualCaptchaChallengeStore.GetOrCreate(HttpContext, RegisterCaptchaScope);
            return;
        }

        ManualCaptchaChallengeStore.Clear(HttpContext, RegisterCaptchaScope);
        CaptchaManualCodigo = string.Empty;
    }
}
