using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class AyudaController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId, string? modulo)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        // Reutiliza permisos del dashboard para habilitar acceso a la ayuda operativa.
        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "DASHBOARD");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        ViewData["Title"] = "Ayuda operativa";
        ViewData["ModuloAyuda"] = modulo?.Trim().ToUpperInvariant();
        return View(baseVm);
    }
}
