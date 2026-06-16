using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Panel;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class PanelController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
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

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(empresaId, cancellationToken);

        var model = new PanelIndexViewModel
        {
            IdEmpresa = empresaId,
            NombreEmpresa = currentCompanyAccessor.EmpresaNombre ?? contexto?.RazonSocial ?? "Empresa",
            Ruc = contexto?.Ruc,
            EstadoPanel = ResolverEstadoPanel(contexto),
            ResumenVigencia = ResolverResumenVigencia(contexto),
            Modulos =
            [
                new PanelModuloViewModel
                {
                    Titulo = "Asientos contables",
                    Descripcion = "Registro manual y control del libro diario.",
                    Icono = "bi-journal-check",
                    Url = "#",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Plan de cuentas",
                    Descripcion = "Mantenimiento del catalogo contable por empresa.",
                    Icono = "bi-diagram-3",
                    Url = "#",
                    Habilitado = false
                },
                new PanelModuloViewModel
                {
                    Titulo = "Compras",
                    Descripcion = "Provisiones, proveedores y contabilizacion de compras.",
                    Icono = "bi-cart-check",
                    Url = "#",
                    Habilitado = false
                },
                new PanelModuloViewModel
                {
                    Titulo = "Ventas",
                    Descripcion = "Clientes, comprobantes y asiento automatico de ventas.",
                    Icono = "bi-cash-stack",
                    Url = "#",
                    Habilitado = false
                }
            ]
        };

        return View(model);
    }

    private static string ResolverEstadoPanel(ContextoSuscripcionEmpresaDto? empresa)
    {
        if (empresa is null)
        {
            return "Sin configuracion comercial";
        }

        if (!empresa.Activo || !empresa.EstadoEmpresa || !empresa.EstadoCuenta)
        {
            return "Cuenta o empresa inactiva";
        }

        return empresa.EstadoSuscripcion switch
        {
            "ACTIVO" => "Contrato activo",
            "TRIAL" => "Periodo de prueba",
            "SUSPENDIDO" => "Suscripcion suspendida",
            "BAJA" => "Servicio de baja",
            _ => "Estado pendiente"
        };
    }

    private static string ResolverResumenVigencia(ContextoSuscripcionEmpresaDto? empresa)
    {
        if (empresa is null)
        {
            return "Aun no hay suscripcion configurada.";
        }

        if (empresa.EsPrueba && empresa.FechaFinPrueba.HasValue)
        {
            return $"Prueba activa hasta {empresa.FechaFinPrueba:dd/MM/yyyy}.";
        }

        if (empresa.FechaFinPlan.HasValue)
        {
            return $"Contrato vigente hasta {empresa.FechaFinPlan:dd/MM/yyyy}.";
        }

        return "Sin fecha de vigencia registrada.";
    }
}
