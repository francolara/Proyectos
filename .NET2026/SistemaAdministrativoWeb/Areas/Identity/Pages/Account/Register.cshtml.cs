using System.ComponentModel.DataAnnotations;
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
    IOptions<IdentityBehaviorSettings> identityBehaviorOptions,
    ILogger<RegisterModel> logger) : PageModel
{
    [BindProperty]
    public UsuarioRegistroInput Usuario { get; set; } = new();

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
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        TurnstileSiteKey = turnstileOptions.Value.SiteKey;
    }

    public async Task<IActionResult> OnPostUsuarioAsync(string? returnUrl = null)
    {
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
        return error.Code switch
        {
            "DuplicateUserName" => "Ya existe una cuenta registrada con este correo.",
            "DuplicateEmail" => "Ya existe una cuenta registrada con este correo.",
            "InvalidEmail" => "Ingrese un correo electronico valido.",
            "PasswordTooShort" => "La contrasena debe tener al menos 6 caracteres.",
            "PasswordRequiresNonAlphanumeric" => "La contrasena debe incluir al menos un caracter especial.",
            "PasswordRequiresDigit" => "La contrasena debe incluir al menos un numero.",
            "PasswordRequiresLower" => "La contrasena debe incluir al menos una letra minuscula.",
            "PasswordRequiresUpper" => "La contrasena debe incluir al menos una letra mayuscula.",
            _ => error.Description
        };
    }

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
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
