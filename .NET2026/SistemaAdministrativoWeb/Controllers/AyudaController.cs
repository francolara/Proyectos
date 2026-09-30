using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Ayuda;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("AYUDA")]
[AllowRestrictedSubscription]
public class AyudaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ISubscriptionAccessService subscriptionAccessService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(string? modulo, CancellationToken cancellationToken)
    {
        if (User.IsInRole("SuperAdmin") && !currentCompanyAccessor.TieneEmpresaActiva)
        {
            return RedirectToAction("Index", "Plataforma");
        }

        var subscriptionAccess = await subscriptionAccessService.EvaluateAsync(User, cancellationToken);
        if ((!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
            && !subscriptionAccess.IsRestricted)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        return View(AyudaCatalogoFactory.Crear(modulo));
    }

    // Firma: FRANCO LARA - 29/09/2026 | Publica el mapa operativo contable interactivo desde el módulo de Ayuda.
    [HttpGet]
    public IActionResult MapaOperativo()
    {
        ViewData["AdminShell"] = true;
        return View();
    }

    [HttpGet]
    public IActionResult ContenidoMapaOperativo()
    {
        var rutaMapa = Path.Combine(
            Directory.GetCurrentDirectory(),
            ".archify",
            "mapa-operativo-contable-20260929-1645",
            "mapa-operativo-administrador.html");

        return System.IO.File.Exists(rutaMapa)
            ? PhysicalFile(rutaMapa, "text/html")
            : NotFound("No se encontró el archivo del mapa operativo contable.");
    }
}
