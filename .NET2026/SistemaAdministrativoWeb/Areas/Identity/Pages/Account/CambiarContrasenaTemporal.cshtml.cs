using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using SistemaAdministrativoWeb.Infrastructure.Security;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[Authorize]
public class CambiarContrasenaTemporalModel(
    UserManager<IdentityUser> userManager,
    SignInManager<IdentityUser> signInManager,
    ILogger<CambiarContrasenaTemporalModel> logger) : PageModel
{
    [BindProperty]
    public InputModel Input { get; set; } = new();

    [BindProperty(SupportsGet = true)]
    public string ReturnUrl { get; set; } = string.Empty;

    public sealed class InputModel
    {
        [Required(ErrorMessage = "Ingrese la contrasena temporal actual.")]
        [DataType(DataType.Password)]
        public string PasswordActual { get; set; } = string.Empty;

        [Required(ErrorMessage = "Ingrese la nueva contrasena.")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "La contrasena debe tener entre 6 y 100 caracteres.")]
        [DataType(DataType.Password)]
        public string NuevaPassword { get; set; } = string.Empty;

        [Required(ErrorMessage = "Confirme la nueva contrasena.")]
        [Compare(nameof(NuevaPassword), ErrorMessage = "Las contrasenas no coinciden.")]
        [DataType(DataType.Password)]
        public string ConfirmarPassword { get; set; } = string.Empty;
    }

    public async Task<IActionResult> OnGetAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        var user = await userManager.GetUserAsync(User);
        if (user is null || !await RequiereCambioContrasenaTemporalAsync(user))
        {
            return LocalRedirect(Url.Content("~/"));
        }

        if (!UsuarioYaVerificado(user))
        {
            return RedirectToPage("./VerificacionTemporal", new { returnUrl = ReturnUrl });
        }

        return Page();
    }

    public async Task<IActionResult> OnPostAsync(string? returnUrl = null)
    {
        ReturnUrl = returnUrl ?? Url.Content("~/");
        var user = await userManager.GetUserAsync(User);
        if (user is null || !await RequiereCambioContrasenaTemporalAsync(user))
        {
            return LocalRedirect(Url.Content("~/"));
        }

        if (!UsuarioYaVerificado(user))
        {
            return RedirectToPage("./VerificacionTemporal", new { returnUrl = ReturnUrl });
        }

        if (!ModelState.IsValid)
        {
            return Page();
        }

        var changeResult = await userManager.ChangePasswordAsync(user, Input.PasswordActual, Input.NuevaPassword);
        if (!changeResult.Succeeded)
        {
            foreach (var error in changeResult.Errors)
            {
                ModelState.AddModelError(string.Empty, IdentityErrorTranslator.Translate(error));
            }

            return Page();
        }

        var claims = await userManager.GetClaimsAsync(user);
        var temporaryClaims = claims
            .Where(x => string.Equals(x.Type, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType, StringComparison.Ordinal))
            .ToList();

        if (temporaryClaims.Count > 0)
        {
            await userManager.RemoveClaimsAsync(user, temporaryClaims);
        }

        HttpContext.Session.Remove(TemporaryPasswordFlowConstants.VerificationSessionKey);
        await signInManager.RefreshSignInAsync(user);
        logger.LogInformation("Usuario completo el cambio obligatorio de contrasena temporal.");
        return LocalRedirect(ReturnUrl);
    }

    private bool UsuarioYaVerificado(IdentityUser user)
        => string.Equals(
            HttpContext.Session.GetString(TemporaryPasswordFlowConstants.VerificationSessionKey),
            user.Id,
            StringComparison.Ordinal);

    private async Task<bool> RequiereCambioContrasenaTemporalAsync(IdentityUser user)
    {
        var claims = await userManager.GetClaimsAsync(user);
        return claims.Any(x =>
            string.Equals(x.Type, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimType, StringComparison.Ordinal)
            && string.Equals(x.Value, TemporaryPasswordFlowConstants.RequirePasswordChangeClaimValue, StringComparison.Ordinal));
    }
}
