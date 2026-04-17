using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class LoginModel(
    SignInManager<ApplicationUser> signInManager,
    UserManager<ApplicationUser> userManager,
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
        var user = await userManager.FindByEmailAsync(email);
        if (user is null)
        {
            ModelState.AddModelError(string.Empty, "Intento de inicio de sesion no valido.");
            await CargarBannerLateralAsync();
            return Page();
        }

        var result = await signInManager.PasswordSignInAsync(user.UserName ?? email, Input.Password, Input.RememberMe, lockoutOnFailure: false);
        if (result.Succeeded)
        {
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
}
