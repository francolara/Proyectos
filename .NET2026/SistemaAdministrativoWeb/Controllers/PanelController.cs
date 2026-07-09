using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Panel;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class PanelController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    IPeriodoContableService periodoContableService,
    ICompraRepository compraRepository,
    IVentaRepository ventaRepository,
    IAsientoRepository asientoRepository,
    ICuentaCorrienteRepository cuentaCorrienteRepository,
    ICajaBancoRepository cajaBancoRepository,
    IAplicacionNotaCreditoRepository aplicacionNotaCreditoRepository) : Controller
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
        var (anioActual, mesActual) = NormalizarPeriodo();
        var periodoActual = $"{anioActual:0000}{mesActual:00}";

        var estadoPeriodoTask = periodoContableService.ObtenerEstadoAsync(empresaId, anioActual, mesActual, cancellationToken);
        var comprasPeriodoTask = compraRepository.ListarPorEmpresaAsync(empresaId, periodoActual, cancellationToken);
        var ventasPeriodoTask = ventaRepository.ListarPorEmpresaAsync(empresaId, periodoActual, cancellationToken);
        var asientosPeriodoTask = asientoRepository.ListarPorEmpresaAsync(empresaId, periodoActual, false, cancellationToken);
        var cuentasCorrientesTask = cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken);
        var movimientosBancariosPeriodoTask = cajaBancoRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, anioActual, mesActual, null, 1, int.MaxValue, cancellationToken);
        var aplicacionesPeriodoTask = aplicacionNotaCreditoRepository.ListarPaginadoPorEmpresaAsync(empresaId, anioActual, mesActual, null, 1, int.MaxValue, cancellationToken);

        await Task.WhenAll(estadoPeriodoTask, comprasPeriodoTask, ventasPeriodoTask, asientosPeriodoTask, cuentasCorrientesTask, movimientosBancariosPeriodoTask, aplicacionesPeriodoTask);

        var estadoPeriodo = await estadoPeriodoTask;
        var comprasPeriodo = (await comprasPeriodoTask).ToList();
        var ventasPeriodo = (await ventasPeriodoTask).ToList();
        var asientosPeriodo = (await asientosPeriodoTask).ToList();
        var cuentasCorrientes = (await cuentasCorrientesTask).ToList();
        var movimientosBancariosPeriodo = await movimientosBancariosPeriodoTask;
        var aplicacionesPeriodo = await aplicacionesPeriodoTask;

        var periodosHistoricos = ConstruirPeriodosHistoricos(6);
        var comprasHistoricasTasks = periodosHistoricos
            .Select(periodo => compraRepository.ListarPorEmpresaAsync(empresaId, periodo.codigo, cancellationToken))
            .ToArray();
        var ventasHistoricasTasks = periodosHistoricos
            .Select(periodo => ventaRepository.ListarPorEmpresaAsync(empresaId, periodo.codigo, cancellationToken))
            .ToArray();

        await Task.WhenAll(comprasHistoricasTasks.Cast<Task>().Concat(ventasHistoricasTasks.Cast<Task>()));

        var comprasPorPeriodo = periodosHistoricos
            .Zip(comprasHistoricasTasks, (periodo, task) => new PanelSeriePeriodoViewModel
            {
                Periodo = periodo.etiqueta,
                Registros = task.Result.Count,
                ImporteTotal = task.Result.Sum(x => x.ImporteTotal)
            })
            .ToList();
        var ventasPorPeriodo = periodosHistoricos
            .Zip(ventasHistoricasTasks, (periodo, task) => new PanelSeriePeriodoViewModel
            {
                Periodo = periodo.etiqueta,
                Registros = task.Result.Count,
                ImporteTotal = task.Result.Sum(x => x.ImporteTotal)
            })
            .ToList();

        var distribucionRegistros = ConstruirDistribucionRegistros(
            comprasPeriodo.Count,
            ventasPeriodo.Count,
            asientosPeriodo.Count,
            movimientosBancariosPeriodo.TotalRecords,
            aplicacionesPeriodo.TotalRecords);

        var model = new PanelIndexViewModel
        {
            IdEmpresa = empresaId,
            NombreEmpresa = currentCompanyAccessor.EmpresaNombre ?? contexto?.RazonSocial ?? "Empresa",
            Ruc = contexto?.Ruc,
            EstadoPanel = ResolverEstadoPanel(contexto),
            ResumenVigencia = ResolverResumenVigencia(contexto),
            PeriodoActual = periodoActual,
            PeriodoCerrado = estadoPeriodo.Cerrado,
            EstadoPeriodo = estadoPeriodo.Cerrado ? "Periodo cerrado" : "Periodo abierto",
            TotalComprasPeriodo = comprasPeriodo.Count,
            ImporteComprasPeriodo = comprasPeriodo.Sum(x => x.ImporteTotal),
            TotalVentasPeriodo = ventasPeriodo.Count,
            ImporteVentasPeriodo = ventasPeriodo.Sum(x => x.ImporteTotal),
            TotalAsientosPeriodo = asientosPeriodo.Count,
            ImporteAsientosPeriodo = asientosPeriodo.Sum(x => x.TotalImporteS),
            TotalCuentasCorrientesActivas = cuentasCorrientes.Count,
            ComprasPorPeriodo = comprasPorPeriodo,
            VentasPorPeriodo = ventasPorPeriodo,
            DistribucionRegistros = distribucionRegistros,
            Indicadores =
            [
                new PanelIndicadorViewModel
                {
                    Titulo = "Compras pendientes",
                    Valor = comprasPeriodo.Count(x => x.Saldo > 0).ToString("0"),
                    Descripcion = "Comprobantes de compra con saldo por pagar en el periodo actual.",
                    Icono = "bi-hourglass-split"
                },
                new PanelIndicadorViewModel
                {
                    Titulo = "Ventas pendientes",
                    Valor = ventasPeriodo.Count(x => x.Saldo > 0).ToString("0"),
                    Descripcion = "Comprobantes de venta con saldo por cobrar en el periodo actual.",
                    Icono = "bi-cash-coin"
                },
                new PanelIndicadorViewModel
                {
                    Titulo = "Compras sin asiento",
                    Valor = comprasPeriodo.Count(x => !x.IdAsiento.HasValue).ToString("0"),
                    Descripcion = "Registros de compras aun no vinculados a un asiento contable.",
                    Icono = "bi-journal-x"
                },
                new PanelIndicadorViewModel
                {
                    Titulo = "CPE validados",
                    Valor = comprasPeriodo.Count(x => x.FechaValidacionCpe.HasValue).ToString("0"),
                    Descripcion = "Compras del periodo con validacion CPE ya registrada.",
                    Icono = "bi-patch-check"
                }
            ],
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
                    Descripcion = "Seguimiento contable por cuenta, auxiliar y documento para revision operativa.",
                    Icono = "bi-bar-chart",
                    Url = "/Reporte/AnalisisCuentas",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Libro Diario",
                    Descripcion = "Consulta de asientos del periodo en vistas Diario auxiliar, Por Cuenta y Por Origen.",
                    Icono = "bi-journal-richtext",
                    Url = "/Reporte/LibroDiario",
                    Habilitado = true
                },
                new PanelModuloViewModel
                {
                    Titulo = "Libro Mayor",
                    Descripcion = "Mayor contable por cuenta con saldo inicial, movimientos y cierre del periodo.",
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

    private static (short anio, byte mes) NormalizarPeriodo()
    {
        var today = DateTime.Today;
        return ((short)today.Year, (byte)today.Month);
    }

    private static IReadOnlyCollection<(string codigo, string etiqueta)> ConstruirPeriodosHistoricos(int totalPeriodos)
    {
        var inicio = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-(Math.Max(1, totalPeriodos) - 1));
        var periodos = new List<(string codigo, string etiqueta)>();

        for (var indice = 0; indice < Math.Max(1, totalPeriodos); indice++)
        {
            var fecha = inicio.AddMonths(indice);
            periodos.Add(($"{fecha.Year:0000}{fecha.Month:00}", $"{fecha.Month:00}/{fecha.Year:0000}"));
        }

        return periodos;
    }

    private static IReadOnlyCollection<PanelDistribucionItemViewModel> ConstruirDistribucionRegistros(int compras, int ventas, int asientos, int movimientosBancarios, int aplicaciones)
    {
        var items = new List<PanelDistribucionItemViewModel>
        {
            new()
            {
                Titulo = "Compras",
                Valor = compras,
                Color = "#2f8cff",
                Descripcion = "Registros de compras del periodo."
            },
            new()
            {
                Titulo = "Ventas",
                Valor = ventas,
                Color = "#1fa971",
                Descripcion = "Registros de ventas del periodo."
            },
            new()
            {
                Titulo = "Asientos",
                Valor = asientos,
                Color = "#f59e0b",
                Descripcion = "Asientos contables emitidos en el periodo."
            },
            new()
            {
                Titulo = "Mov. bancarios",
                Valor = movimientosBancarios,
                Color = "#8b5cf6",
                Descripcion = "Movimientos de caja y bancos del periodo."
            },
            new()
            {
                Titulo = "Aplicaciones",
                Valor = aplicaciones,
                Color = "#ef4444",
                Descripcion = "Aplicaciones de notas de credito registradas en el periodo."
            }
        };

        var total = items.Sum(x => x.Valor);
        if (total <= 0)
        {
            return items.Select(x => new PanelDistribucionItemViewModel
            {
                Titulo = x.Titulo,
                Valor = x.Valor,
                Color = x.Color,
                Descripcion = x.Descripcion,
                Porcentaje = 0m
            }).ToList();
        }

        return items.Select(x => new PanelDistribucionItemViewModel
        {
            Titulo = x.Titulo,
            Valor = x.Valor,
            Color = x.Color,
            Descripcion = x.Descripcion,
            Porcentaje = Math.Round((decimal)x.Valor * 100m / total, 2)
        }).ToList();
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
