using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Models;
using System.Diagnostics;

namespace SistemaAdministrativoWeb.Controllers;

public class HomeController(ICurrentCompanyAccessor currentCompanyAccessor) : Controller
{
    public IActionResult Index()
    {
        if (User.Identity?.IsAuthenticated != true)
        {
            return Redirect("/Identity/Account/Login");
        }

        if (User.IsInRole("SuperAdmin"))
        {
            return RedirectToAction("Index", "Plataforma");
        }

        if (!currentCompanyAccessor.TieneEmpresaActiva)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        return RedirectToAction("Index", "Panel");
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
