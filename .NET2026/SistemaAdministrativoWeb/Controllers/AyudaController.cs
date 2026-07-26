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
}
