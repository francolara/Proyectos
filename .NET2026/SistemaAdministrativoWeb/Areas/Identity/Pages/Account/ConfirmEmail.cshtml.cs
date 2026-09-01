using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaAdministrativoWeb.Infrastructure.Email;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ConfirmEmailModel(
    UserManager<IdentityUser> userManager,
    IAccountEmailService accountEmailService,
    ILogger<ConfirmEmailModel> logger) : PageModel
{
    // Firma: FRANCO LARA - 31/08/2026 | Confirma la cuenta y envia el correo HTML de bienvenida.
    public string StatusMessage { get; private set; } = string.Empty;
    public bool IsSuccess { get; private set; }
    public string LoginUrl { get; private set; } = "/Identity/Account/Login";

    public async Task<IActionResult> OnGetAsync(string? userId, string? code, string? returnUrl = null)
    {
        LoginUrl = Url.Page(
                "/Account/Login",
                pageHandler: null,
                values: new { area = "Identity", returnUrl },
                protocol: Request.Scheme)
            ?? "/Identity/Account/Login";

        if (string.IsNullOrWhiteSpace(userId) || string.IsNullOrWhiteSpace(code))
        {
            SetInvalidStatus();
            return Page();
        }

        var user = await userManager.FindByIdAsync(userId);
        if (user is null)
        {
            SetInvalidStatus();
            return Page();
        }

        string decodedCode;
        try
        {
            decodedCode = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));
        }
        catch (FormatException)
        {
            SetInvalidStatus();
            return Page();
        }

        var result = await userManager.ConfirmEmailAsync(user, decodedCode);
        if (!result.Succeeded)
        {
            SetInvalidStatus();
            return Page();
        }

        IsSuccess = true;
        StatusMessage = "Correo confirmado correctamente. Tu cuenta ya esta activa.";

        if (accountEmailService.IsEnabled && !string.IsNullOrWhiteSpace(user.Email))
        {
            try
            {
                await accountEmailService.SendWelcomeEmailAsync(
                    user.Email,
                    user.Email,
                    LoginUrl,
                    HttpContext.RequestAborted);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "La cuenta {UserId} se confirmo, pero no se pudo enviar la bienvenida.", user.Id);
            }
        }

        return Page();
    }

    private void SetInvalidStatus()
    {
        IsSuccess = false;
        StatusMessage = "El enlace de confirmacion no es valido o ya fue utilizado. Solicita uno nuevo desde el inicio de sesion.";
    }
}
