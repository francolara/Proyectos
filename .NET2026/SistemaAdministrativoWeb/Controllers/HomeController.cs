using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.Models;
using System.Diagnostics;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Controllers;

public class HomeController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IModulePermissionService modulePermissionService,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository) : Controller
{
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        if (User.Identity?.IsAuthenticated != true)
        {
            return View();
        }

        if (User.IsInRole("SuperAdmin"))
        {
            return RedirectToAction("Index", "Plataforma");
        }

        if (!currentCompanyAccessor.TieneEmpresaActiva)
        {
            var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrWhiteSpace(aspNetUserId))
            {
                var contextoLogin = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(aspNetUserId, cancellationToken);
                if (contextoLogin is null || !contextoLogin.TieneAcceso)
                {
                    return RedirectToAction("Index", "EmpresaContexto");
                }

                if (contextoLogin?.TieneAcceso == true)
                {
                    if (contextoLogin.SoloModulosCuenta)
                    {
                        if (await modulePermissionService.CanAccessModuleAsync(User, "CONFIGURACION", cancellationToken))
                        {
                            return RedirectToAction("Index", "Configuracion");
                        }

                        if (await modulePermissionService.CanAccessModuleAsync(User, "MISUSCRIPCION", cancellationToken))
                        {
                            return RedirectToAction("Index", "MiSuscripcion");
                        }

                        if (await modulePermissionService.CanAccessModuleAsync(User, "AYUDA", cancellationToken))
                        {
                            return RedirectToAction("Index", "Ayuda");
                        }
                    }

                    if (await modulePermissionService.CanAccessModuleAsync(User, "EMPRESAS", cancellationToken))
                    {
                        return RedirectToAction("Index", "EmpresaContexto");
                    }

                    if (await modulePermissionService.CanAccessModuleAsync(User, "CONFIGURACION", cancellationToken))
                    {
                        return RedirectToAction("Index", "Configuracion");
                    }

                    if (await modulePermissionService.CanAccessModuleAsync(User, "MISUSCRIPCION", cancellationToken))
                    {
                        return RedirectToAction("Index", "MiSuscripcion");
                    }

                    if (await modulePermissionService.CanAccessModuleAsync(User, "AYUDA", cancellationToken))
                    {
                        return RedirectToAction("Index", "Ayuda");
                    }
                }
            }

            TempData["ErrorMessage"] = "No se encontro una opcion inicial permitida para este usuario.";
            return View();
        }

        if (!await modulePermissionService.CanAccessModuleAsync(User, "DASHBOARD", cancellationToken))
        {
            if (await modulePermissionService.CanAccessModuleAsync(User, "MISUSCRIPCION", cancellationToken))
            {
                return RedirectToAction("Index", "MiSuscripcion");
            }

            if (await modulePermissionService.CanAccessModuleAsync(User, "AYUDA", cancellationToken))
            {
                return RedirectToAction("Index", "Ayuda");
            }

            TempData["ErrorMessage"] = "No tiene acceso al dashboard de la empresa activa.";
            return View();
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
