using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaAdministrativoWeb.Infrastructure.Security;

namespace SistemaAdministrativoWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ResetPasswordModel(UserManager<IdentityUser> userManager) : PageModel
{
    // Firma: FRANCO LARA - 31/08/2026 | Restablece la contrasena mediante token seguro de ASP.NET Identity.
    [BindProperty]
    public InputModel Input { get; set; } = new();

    public sealed class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingrese un correo valido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "La nueva contrasena es obligatoria.")]
        [StringLength(100, MinimumLength = 6, ErrorMessage = "La contrasena debe tener entre 6 y 100 caracteres.")]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [Required(ErrorMessage = "Confirme la nueva contrasena.")]
        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "Las contrasenas no coinciden.")]
        public string ConfirmPassword { get; set; } = string.Empty;

        [Required]
        public string Code { get; set; } = string.Empty;
    }

    public IActionResult OnGet(string? code = null, string? email = null)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return BadRequest("El enlace de recuperacion no es valido.");
        }

        Input = new InputModel { Code = code, Email = email ?? string.Empty };
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            return Page();
        }

        var user = await userManager.FindByEmailAsync(Input.Email.Trim());
        if (user is null)
        {
            return RedirectToPage("./ResetPasswordConfirmation");
        }

        string code;
        try
        {
            code = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(Input.Code));
        }
        catch (FormatException)
        {
            ModelState.AddModelError(string.Empty, "El enlace de recuperacion no es valido. Solicita uno nuevo.");
            return Page();
        }

        var result = await userManager.ResetPasswordAsync(user, code, Input.Password);
        if (result.Succeeded)
        {
            return RedirectToPage("./ResetPasswordConfirmation");
        }

        foreach (var error in result.Errors)
        {
            ModelState.AddModelError(string.Empty, IdentityErrorTranslator.Translate(error));
        }

        return Page();
    }
}
