using GestionServicios.Web.Navigation;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
[Authorize]
public sealed class ModuleController(IAuthorizationService authorization) : Controller
{
    [HttpGet("/modulo/{code}")]
    public async Task<IActionResult> Index(string code)
    {
        var entry = HytMenu.Items.FirstOrDefault(item => string.Equals(item.Code, code, StringComparison.OrdinalIgnoreCase));
        if (entry is null) return NotFound();
        if (!(await authorization.AuthorizeAsync(User, null, "perm:" + entry.Permission)).Succeeded)
        { Response.StatusCode = 403; return View("~/Views/Shared/Denied.cshtml"); }
        if (GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find(code) is not null)
            return RedirectToAction("Index", "Master", new { entity = code });
        if (code == "configuracion") return RedirectToAction("Index", "Company");
        if (code == "presupuestos") return RedirectToAction("Index", "Budget");
        if (code is "oc-cliente" or "guias") return RedirectToAction("Index", "Followup", new { kind = code });
        if (code == "reportes") return RedirectToAction("Index", "Reports");
        return View(entry);
    }
}
