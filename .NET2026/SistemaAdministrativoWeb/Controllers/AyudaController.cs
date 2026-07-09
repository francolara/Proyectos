using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Ayuda;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class AyudaController(ICurrentCompanyAccessor currentCompanyAccessor) : Controller
{
    [HttpGet]
    public IActionResult Index(string? modulo)
    {
        if (User.IsInRole("SuperAdmin") && !currentCompanyAccessor.TieneEmpresaActiva)
        {
            return RedirectToAction("Index", "Plataforma");
        }

        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        return View(AyudaCatalogoFactory.Crear(modulo));
    }
}
