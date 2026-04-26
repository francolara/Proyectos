using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ResetPasswordModel(UserManager<ApplicationUser> userManager) : PageModel
{
    [BindProperty]
    public InputModel Input { get; set; } = new();

    public class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingresa un correo valido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "La contrasena es obligatoria.")]
        [StringLength(100, ErrorMessage = "La contrasena debe tener al menos {2} y como maximo {1} caracteres.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        public string Password { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "La contrasena y la confirmacion no coinciden.")]
        [Required(ErrorMessage = "La confirmacion de contrasena es obligatoria.")]
        public string ConfirmPassword { get; set; } = string.Empty;

        [Required]
        public string Code { get; set; } = string.Empty;
    }

    public IActionResult OnGet(string? code = null, string? email = null)
    {
        if (string.IsNullOrWhiteSpace(code))
        {
            return BadRequest("Se requiere un codigo de restablecimiento.");
        }

        Input = new InputModel
        {
            Code = code,
            Email = email ?? string.Empty
        };
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid)
        {
            return Page();
        }

        var user = await userManager.FindByEmailAsync((Input.Email ?? string.Empty).Trim());
        if (user is null)
        {
            return RedirectToPage("./ResetPasswordConfirmation");
        }

        var code = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(Input.Code));
        var result = await userManager.ResetPasswordAsync(user, code, Input.Password);
        if (result.Succeeded)
        {
            return RedirectToPage("./ResetPasswordConfirmation");
        }

        foreach (var error in result.Errors)
        {
            ModelState.AddModelError(string.Empty, TraducirErrorIdentity(error.Code));
        }

        return Page();
    }

    private static string TraducirErrorIdentity(string code)
    {
        return code switch
        {
            "PasswordRequiresNonAlphanumeric" => "La contrasena debe incluir al menos un simbolo.",
            "PasswordRequiresLower" => "La contrasena debe incluir al menos una letra minuscula.",
            "PasswordRequiresUpper" => "La contrasena debe incluir al menos una letra mayuscula.",
            "PasswordRequiresDigit" => "La contrasena debe incluir al menos un numero.",
            "PasswordTooShort" => "La contrasena es muy corta.",
            "InvalidToken" => "El enlace de recuperacion ya no es valido. Solicita uno nuevo.",
            _ => "No se pudo restablecer la contrasena. Intenta nuevamente."
        };
    }
}
