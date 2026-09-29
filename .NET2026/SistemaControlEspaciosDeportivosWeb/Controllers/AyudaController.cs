using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using SistemaControlEspaciosDeportivosWeb.ViewModels.Ayuda;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class AyudaController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    IWebHostEnvironment webHostEnvironment)
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
        return View(AyudaCatalogoFactory.Crear(baseVm, modulo));
    }

    public async Task<IActionResult> MapaOperativo(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "DASHBOARD");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        ViewData["Title"] = "Mapa operativo";
        return View(baseVm);
    }

    public async Task<IActionResult> ContenidoMapaOperativo(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "DASHBOARD");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return Forbid();

        var rutaMapa = Path.Combine(
            webHostEnvironment.ContentRootPath,
            ".archify",
            "mapa-operativo-administrador-20260929-0018",
            "guia-operativa-interactiva.html");

        if (!System.IO.File.Exists(rutaMapa))
            return NotFound("No se encontró el archivo del mapa operativo.");

        Response.Headers.CacheControl = "no-store";
        return PhysicalFile(rutaMapa, "text/html; charset=utf-8");
    }
}
