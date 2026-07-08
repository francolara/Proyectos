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
                    Url = "/Asiento",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Plan de cuentas",
                    Descripcion = "Mantenimiento del catalogo contable por empresa.",
                    Icono = "bi-diagram-3",
                    Url = "/PlanCuenta",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Personas",
                    Descripcion = "Mantenimiento de clientes, proveedores y terceros por empresa.",
                    Icono = "bi-people",
                    Url = "/Persona",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Origenes contables",
                    Descripcion = "Configuracion de subdiarios y circuitos de registro por empresa.",
                    Icono = "bi-journal-richtext",
                    Url = "/Origen",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Cuentas destino",
                    Descripcion = "Equivalencias contables por ejercicio para distribuciones automaticas.",
                    Icono = "bi-sliders2",
                    Url = "/CuentaDestinoRegla",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Config. compras/ventas",
                    Descripcion = "Parametros contables por escenario para asiento automatico de compras y ventas.",
                    Icono = "bi-gear-wide-connected",
                    Url = "/ConfiguracionContabilizacion",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Compras",
                    Descripcion = "Provisiones, proveedores y contabilizacion de compras.",
                    Icono = "bi-cart-check",
                    Url = "/Compra",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Ventas",
                    Descripcion = "Clientes, comprobantes y asiento automatico de ventas.",
                    Icono = "bi-cash-stack",
                    Url = "/Venta",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Aplicaciones",
                    Descripcion = "Compensacion parcial o total de comprobantes con notas de credito y asiento APNC.",
                    Icono = "bi-link-45deg",
                    Url = "/Aplicacion",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Analisis de cuentas",
                    Descripcion = "Reporte HTML de seguimiento por cuenta, documento y auxiliar rescatado del flujo legacy.",
                    Icono = "bi-bar-chart",
                    Url = "/Reporte/AnalisisCuentas",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Libro Diario",
                    Descripcion = "Reporte HTML que replica diario auxiliar y diario por origen del sistema legacy.",
                    Icono = "bi-journal-richtext",
                    Url = "/Reporte/LibroDiario",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Libro Mayor",
                    Descripcion = "Reporte HTML del mayor contable con saldo inicial por cuenta y filtro por Nro documento.",
                    Icono = "bi-journals",
                    Url = "/Reporte/LibroMayor",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Registro de ventas",
                    Descripcion = "Reporte HTML del registro de ventas mensual en formato A4 basado en la provision de ventas.",
                    Icono = "bi-receipt-cutoff",
                    Url = "/Reporte/RegistroVentas",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Registro de compras",
                    Descripcion = "Reporte HTML del registro de compras mensual en formato A4 basado en la provision de compras.",
                    Icono = "bi-cart3",
                    Url = "/Reporte/RegistroCompras",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Balance de comprobacion",
                    Descripcion = "Reporte HTML que replica el balance de comprobacion mensual del legacy por rango de periodos y grado.",
                    Icono = "bi-table",
                    Url = "/Reporte/BalanceComprobacion",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Libros Electronicos",
                    Descripcion = "Generacion, validacion y descarga TXT de PLE 5.1, 5.2 y 6.1 basada en asientos contables existentes.",
                    Icono = "bi-filetype-txt",
                    Url = "/LibroElectronico",
                    Habilitado = true
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
