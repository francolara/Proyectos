using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;

namespace SistemaControlEspaciosDeportivosWeb.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ForgotPasswordModel(
    UserManager<ApplicationUser> userManager,
    IAccountEmailService accountEmailService,
    ILogger<ForgotPasswordModel> logger) : PageModel
{
    [BindProperty]
    public InputModel Input { get; set; } = new();

    public class InputModel
    {
        [Required(ErrorMessage = "El correo es obligatorio.")]
        [EmailAddress(ErrorMessage = "Ingresa un correo valido.")]
        public string Email { get; set; } = string.Empty;
    }

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid) return Page();

        var email = (Input.Email ?? string.Empty).Trim();
        var user = await userManager.FindByEmailAsync(email);

        if (user is not null)
        {
            if (await userManager.IsEmailConfirmedAsync(user))
            {
                var code = await userManager.GeneratePasswordResetTokenAsync(user);
                code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
                var callbackUrl = Url.Page(
                    "/Account/ResetPassword",
                    pageHandler: null,
                    values: new { area = "Identity", code, email },
                    protocol: Request.Scheme);

                if (!string.IsNullOrWhiteSpace(callbackUrl))
                {
                    try
                    {
                        await accountEmailService.SendResetPasswordEmailAsync(email, user.Nombres, callbackUrl);
                    }
                    catch (EmailDeliveryException ex)
                    {
                        logger.LogWarning(ex, "No se pudo enviar correo de recuperacion para {Email}.", email);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error no controlado al enviar correo de recuperacion para {Email}.", email);
                    }
                }
            }
            else
            {
                var code = await userManager.GenerateEmailConfirmationTokenAsync(user);
                code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
                var callbackUrl = Url.Page(
                    "/Account/ConfirmEmail",
                    pageHandler: null,
                    values: new { area = "Identity", userId = user.Id, code },
                    protocol: Request.Scheme);

                if (!string.IsNullOrWhiteSpace(callbackUrl))
                {
                    try
                    {
                        await accountEmailService.SendConfirmationEmailAsync(email, user.Nombres, callbackUrl);
                    }
                    catch (EmailDeliveryException ex)
                    {
                        logger.LogWarning(ex, "No se pudo reenviar confirmacion desde forgot password para {Email}.", email);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError(ex, "Error no controlado al reenviar confirmacion desde forgot password para {Email}.", email);
                    }
                }
            }
        }

        return RedirectToPage("./ForgotPasswordConfirmation");
    }
}
