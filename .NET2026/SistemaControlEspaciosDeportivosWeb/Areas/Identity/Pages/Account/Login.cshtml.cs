using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

// Firma: Codex - 20/04/2026 | Login con expiracion deslizante: 30 minutos sin Recordarme y 2 dias con Recordarme.
[AllowAnonymous]
public class LoginModel(
    SignInManager<ApplicationUser> signInManager,
    UserManager<ApplicationUser> userManager,
    IAccountEmailService accountEmailService,
    ISportCenterStoredProcedureService spService,
    ILogger<LoginModel> logger) : PageModel
{
    [BindProperty]
    public InputModel Input { get; set; } = new();

    public IList<AuthenticationScheme> ExternalLogins { get; set; } = new List<AuthenticationScheme>();
    public string ReturnUrl { get; set; } = string.Empty;
    public WebBannerPublicoViewModel? BannerLateral { get; set; }

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
        ReturnUrl = returnUrl ?? Url.Content("~/");
        await CargarBannerLateralAsync();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        if (!ModelState.IsValid)
        {
            await CargarBannerLateralAsync();
            return Page();
        }

        var email = (Input.Email ?? string.Empty).Trim();
        Input.Email = email;
        var user = await userManager.FindByEmailAsync(email);
        if (user is null)
        {
            ModelState.AddModelError(string.Empty, "Intento de inicio de sesion no valido.");
            await CargarBannerLateralAsync();
            return Page();
        }

        if (!await userManager.IsEmailConfirmedAsync(user))
        {
            ModelState.AddModelError(string.Empty, "Tu cuenta aun no esta confirmada. Revisa tu correo o reenvia el enlace de confirmacion.");
            await CargarBannerLateralAsync();
            return Page();
        }

        var result = await signInManager.CheckPasswordSignInAsync(user, Input.Password, lockoutOnFailure: true);
        if (result.Succeeded)
        {
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
            logger.LogWarning("Cuenta bloqueada.");
            return RedirectToPage("./Lockout");
        }

        ModelState.AddModelError(string.Empty, "Intento de inicio de sesion no valido.");
        await CargarBannerLateralAsync();
        return Page();
    }

    public async Task<IActionResult> OnPostResendConfirmationAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        ExternalLogins = (await signInManager.GetExternalAuthenticationSchemesAsync()).ToList();
        ReturnUrl = returnUrl;

        var email = (Input.Email ?? string.Empty).Trim();
        Input.Email = email;
        if (string.IsNullOrWhiteSpace(email))
        {
            ModelState.AddModelError(string.Empty, "Ingresa tu correo para reenviar la confirmacion.");
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
                EmailConfirmed = true,
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
}
