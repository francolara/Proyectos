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
        if (code == "usuarios" && !User.HasClaim(Security.AccessService.PermissionClaim, "USUARIOS.CONSULTAR") &&
            User.HasClaim(Security.AccessService.PermissionClaim, "PERFILES.CONSULTAR"))
            return RedirectToAction("Index", "Profiles");
        if (!(await authorization.AuthorizeAsync(User, null, "perm:" + entry.Permission)).Succeeded)
        { Response.StatusCode = 403; return View("~/Views/Shared/Denied.cshtml"); }
        if (GestionServicios.Application.Maintenance.MaintenanceDefinitions.Find(code) is not null)
            return RedirectToAction("Index", "Master", new { entity = code });
        if (code == "configuracion") return RedirectToAction("Index", "Company");
        if (code == "presupuestos") return RedirectToAction("Index", "Budget");
        if (code is "oc-cliente" or "guias") return RedirectToAction("Index", "Followup", new { kind = code });
        if (code == "reportes") return RedirectToAction("Index", "Reports");
        if (entry.Code == "pagos-pendientes-proveedores") return RedirectToAction("Pending", "SupplierPayments");
        if (entry.Code == "pagos-proveedores") return RedirectToAction("Index", "SupplierPayments");
        if (entry.Code == "proveedores") return RedirectToAction("Index", "Suppliers");
        if (entry.Code == "trabajadores") return RedirectToAction("Index", "Workers");
        if (entry.Code == "planificacion-pagos") return RedirectToAction("Index", "PayPlanning");
        if (entry.Code == "cargos") return RedirectToAction("Index", "WorkerPositions");
        if (entry.Code == "reporte-planillas") return RedirectToAction("Index", "PayrollReports");
        if (entry.Code == "obligaciones-planillas") return RedirectToAction("Index", "PayrollObligations");
        if (entry.Code == "sctr") return RedirectToAction("Sctr", "PayrollObligations");
        if (code == "usuarios") return RedirectToAction("Index", "Users");
        if (code == "auditoria") return RedirectToAction("Index", "Audit");
        return View(entry);
    }
}
