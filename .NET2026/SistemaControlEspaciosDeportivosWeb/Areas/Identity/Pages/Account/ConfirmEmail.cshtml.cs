using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ConfirmEmailModel(UserManager<ApplicationUser> userManager) : PageModel
{
    public string StatusMessage { get; private set; } = string.Empty;
    public bool IsSuccess { get; private set; }
    public string LoginUrl { get; private set; } = "/Identity/Account/Login";

    public async Task<IActionResult> OnGetAsync(string? userId, string? code, string? returnUrl = null)
    {
        LoginUrl = Url.Page("/Account/Login", new { area = "Identity", returnUrl }) ?? "/Identity/Account/Login";

        if (string.IsNullOrWhiteSpace(userId) || string.IsNullOrWhiteSpace(code))
        {
            StatusMessage = "El enlace de confirmacion no es valido.";
            IsSuccess = false;
            return Page();
        }

        var user = await userManager.FindByIdAsync(userId);
        if (user is null)
        {
            StatusMessage = "No se encontro la cuenta asociada al enlace.";
            IsSuccess = false;
            return Page();
        }

        var decodedCode = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));
        var result = await userManager.ConfirmEmailAsync(user, decodedCode);
        if (result.Succeeded)
        {
            StatusMessage = "Correo confirmado correctamente. Ya puedes iniciar sesion.";
            IsSuccess = true;
            return Page();
        }

        StatusMessage = "No se pudo confirmar el correo. Solicita un nuevo enlace.";
        IsSuccess = false;
        return Page();
    }
}
