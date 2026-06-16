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
public class RegisterModel(
    UserManager<IdentityUser> userManager,
    SignInManager<IdentityUser> signInManager,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ITurnstileValidationService turnstileValidationService,
    IOptions<CloudflareTurnstileSettings> turnstileOptions,
    ILogger<RegisterModel> logger) : PageModel
{
    [BindProperty]
    public string TipoRegistro { get; set; } = "usuario";

    [BindProperty]
    public UsuarioRegistroInput Usuario { get; set; } = new();

    [BindProperty]
    public EmpresaRegistroInput Empresa { get; set; } = new();

    public string ReturnUrl { get; set; } = string.Empty;
    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string TurnstileSiteKey { get; private set; } = string.Empty;

    public sealed class UsuarioRegistroInput
    {
        [Required(ErrorMessage = "Ingrese su nombre completo.")]
        [StringLength(180)]
        public string NombreCompleto { get; set; } = string.Empty;

        [StringLength(30)]
        public string? Telefono { get; set; }

        [Required(ErrorMessage = "Ingrese su correo.")]
        [EmailAddress]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese su contrasena.")]
        [StringLength(100, MinimumLength = 6)]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "Las contrasenas no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;
    }

    public sealed class EmpresaRegistroInput
    {
        [Required(ErrorMessage = "Ingrese el nombre del contacto.")]
        [StringLength(180)]
        public string NombreContacto { get; set; } = string.Empty;

        [StringLength(30)]
        public string? Telefono { get; set; }

        [Required(ErrorMessage = "Ingrese el correo del negocio.")]
        [EmailAddress]
        public string Correo { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese la razon social.")]
        [StringLength(200)]
        public string RazonSocial { get; set; } = string.Empty;

        [StringLength(200)]
        public string? NombreComercial { get; set; }

        [Required(ErrorMessage = "Ingrese el RUC.")]
        [StringLength(11, MinimumLength = 11, ErrorMessage = "El RUC debe tener 11 digitos.")]
        public string Ruc { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese la contrasena.")]
        [StringLength(100, MinimumLength = 6)]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "Las contrasenas no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;
    }

    public async Task OnGetAsync(string? returnUrl = null, string? tipoRegistro = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
        if (!string.IsNullOrWhiteSpace(tipoRegistro))
        {
            TipoRegistro = tipoRegistro.Trim().Equals("empresa", StringComparison.OrdinalIgnoreCase)
                || tipoRegistro.Trim().Equals("club", StringComparison.OrdinalIgnoreCase)
                ? "empresa"
                : "usuario";
        }
    }

    public async Task<IActionResult> OnPostUsuarioAsync(string? returnUrl = null)
    {
        TipoRegistro = "usuario";
        ReturnUrl = returnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;

        if (!TryValidateModel(Usuario, nameof(Usuario)))
        {
            return Page();
        }

        if (!await ValidarTurnstileAsync())
        {
            return Page();
        }

        var user = new IdentityUser
        {
            UserName = Usuario.Email.Trim(),
            Email = Usuario.Email.Trim(),
            EmailConfirmed = true
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
        await signInManager.SignInAsync(user, isPersistent: false);
        return LocalRedirect(ReturnUrl);
    }

    public async Task<IActionResult> OnPostEmpresaAsync(string? returnUrl = null)
    {
        TipoRegistro = "empresa";
        ReturnUrl = returnUrl ?? Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;

        if (!TryValidateModel(Empresa, nameof(Empresa)))
        {
            return Page();
        }

        if (!await ValidarTurnstileAsync())
        {
            return Page();
        }

        var email = Empresa.Correo.Trim();
        var user = new IdentityUser
        {
            UserName = email,
            Email = email,
            EmailConfirmed = true
        };

        var result = await userManager.CreateAsync(user, Empresa.Password);
        if (!result.Succeeded)
        {
            AgregarErrores(result);
            return Page();
        }

        await userManager.AddToRoleAsync(user, "AdministradorEmpresa");

        await cuentaAdministradoraRepository.RegistrarCuentaConEmpresaAsync(new RegistroCuentaAdministradoraConEmpresaRequest
        {
            AspNetUserId = user.Id,
            NombreCompleto = Empresa.NombreContacto.Trim(),
            Telefono = LimpiarTelefono(Empresa.Telefono),
            CorreoReferencia = email,
            CodigoCuenta = GenerarCodigoCuenta(Empresa.RazonSocial, email),
            NombreCuenta = Empresa.RazonSocial.Trim(),
            CodigoEmpresa = GenerarCodigoEmpresa(Empresa.RazonSocial, Empresa.Ruc),
            RazonSocial = Empresa.RazonSocial.Trim(),
            NombreComercial = string.IsNullOrWhiteSpace(Empresa.NombreComercial) ? Empresa.RazonSocial.Trim() : Empresa.NombreComercial.Trim(),
            Ruc = Empresa.Ruc.Trim(),
            DiasPrueba = 30,
            UsuarioRegistro = email
        });

        logger.LogInformation("Cuenta administradora registrada con empresa principal.");
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
            ModelState.AddModelError(string.Empty, error.Description);
        }
    }

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
    }

    private static string GenerarCodigoEmpresa(string razonSocial, string ruc)
    {
        var baseCodigo = string.IsNullOrWhiteSpace(ruc)
            ? new string(razonSocial.Where(char.IsLetterOrDigit).Take(8).ToArray()).ToUpperInvariant()
            : ruc.Trim();

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = $"EMP{DateTime.UtcNow:HHmmss}";
        }

        return baseCodigo.Length > 20 ? baseCodigo[..20] : baseCodigo;
    }

    private static string GenerarCodigoCuenta(string nombreCuenta, string correo)
    {
        var baseCodigo = new string(nombreCuenta.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = new string(correo.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();
        }

        if (string.IsNullOrWhiteSpace(baseCodigo))
        {
            baseCodigo = $"CTA{DateTime.UtcNow:HHmmss}";
        }

        return baseCodigo.Length > 20 ? baseCodigo[..20] : baseCodigo;
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

        logger.LogWarning("Turnstile rechazo registro. Errores: {Errores}", string.Join(",", resultado.ErrorCodes ?? Array.Empty<string>()));
        ModelState.AddModelError(string.Empty, "No se pudo validar la verificacion de seguridad. Intenta nuevamente.");
        return false;
    }
}
