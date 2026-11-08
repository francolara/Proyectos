using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Security.Claims;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Security;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 07/11/2026 | Cambio obligatorio o voluntario con renovación de la sesión.
[Authorize, PasswordChangeAllowed, Route("Cuenta/CambiarClave")]
[ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
public sealed class AccountController(IAccessRepository access, ISecurityAdministrationRepository administration,
    IPasswordHasher<AccessAccount> hasher, TimeProvider clock) : Controller
{
    [HttpGet]
    public IActionResult Index() => View(new ChangePasswordInput());

    [HttpPost]
    public async Task<IActionResult> Index(ChangePasswordInput input, CancellationToken ct)
    {
        var account = await access.FindByIdAsync(User.FindFirstValue(ClaimTypes.NameIdentifier)!, ct);
        if (account?.PasswordHash is null || !account.Active ||
            account.SecurityStamp != User.FindFirstValue(AccessService.StampClaim) ||
            account.LockoutEnd > clock.GetUtcNow())
            return Challenge();

        if (!account.MustChangePassword && !Matches(account, input.CurrentPassword))
            ModelState.AddModelError(nameof(input.CurrentPassword), "La contraseña actual es incorrecta.");
        if (ModelState.IsValid && Matches(account, input.NewPassword))
            ModelState.AddModelError(nameof(input.NewPassword), "Elige una contraseña diferente de la actual.");

        if (ModelState.IsValid)
        {
            try
            {
                var hash = hasher.HashPassword(account, input.NewPassword);
                await administration.ChangePasswordAsync(account.Id, account.SecurityStamp, account.PasswordHash,
                    hash, account.MustChangePassword, ct);
                var updated = await access.FindByIdAsync(account.Id, ct);
                if (updated is null || !updated.Active || updated.MustChangePassword || updated.PasswordHash != hash)
                {
                    await HttpContext.SignOutAsync("HytCookie");
                    return LocalRedirect("/Acceso");
                }
                var now = clock.GetUtcNow();
                var properties = new AuthenticationProperties { IsPersistent = false, AllowRefresh = true };
                properties.Items["hyt:created"] = now.UtcTicks.ToString(CultureInfo.InvariantCulture);
                await HttpContext.SignInAsync("HytCookie", AccessService.Principal(updated, now.AddHours(8).UtcTicks), properties);
                TempData["PasswordChanged"] = true;
                return LocalRedirect("/");
            }
            catch (MaintenanceException error) { ModelState.AddModelError("", error.Message); }
        }
        return View(new ChangePasswordInput());
    }

    private bool Matches(AccessAccount account, string? password) =>
        !string.IsNullOrEmpty(password) && password.Length <= 128 &&
        hasher.VerifyHashedPassword(account, account.PasswordHash!, password) != PasswordVerificationResult.Failed;
}

public sealed class ChangePasswordInput
{
    [StringLength(128)]
    public string? CurrentPassword { get; set; }

    [Required(ErrorMessage = "Escribe una contraseña nueva.")]
    [StringLength(128, MinimumLength = 12, ErrorMessage = "Usa entre 12 y 128 caracteres.")]
    public string NewPassword { get; set; } = "";

    [Compare(nameof(NewPassword), ErrorMessage = "Las contraseñas no coinciden.")]
    public string? ConfirmPassword { get; set; }
}
