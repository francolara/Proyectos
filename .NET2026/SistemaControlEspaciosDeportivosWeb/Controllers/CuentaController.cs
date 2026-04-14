using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class CuentaController(UserManager<ApplicationUser> userManager, SignInManager<ApplicationUser> signInManager) : Controller
{
    [HttpGet]
    public async Task<IActionResult> CambiarContrasena()
    {
        var usuario = await userManager.GetUserAsync(User);
        return View(new CambiarContrasenaViewModel
        {
            CorreoUsuario = usuario?.Email ?? User.Identity?.Name ?? string.Empty
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CambiarContrasena(CambiarContrasenaViewModel model)
    {
        model.CorreoUsuario = (await userManager.GetUserAsync(User))?.Email ?? User.Identity?.Name ?? string.Empty;

        if (!ModelState.IsValid)
            return View(model);

        var usuario = await userManager.GetUserAsync(User);
        if (usuario is null)
            return Challenge();

        var resultado = await userManager.ChangePasswordAsync(usuario, model.ContrasenaActual, model.NuevaContrasena);
        if (!resultado.Succeeded)
        {
            foreach (var error in resultado.Errors)
                ModelState.AddModelError(string.Empty, error.Description);
            return View(model);
        }

        await signInManager.RefreshSignInAsync(usuario);
        TempData["CuentaOk"] = "Contrasena actualizada correctamente.";
        return RedirectToAction(nameof(CambiarContrasena));
    }
}
