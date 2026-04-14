using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class PanelController(ISportCenterStoredProcedureService spService, IModuloPermisoService moduloPermisoService) : Controller
{
    public async Task<IActionResult> Index(int? negocioId, int? sedeId, DateOnly? fechaDesde, DateOnly? fechaHasta)
    {
        ViewData["AdminShell"] = true;

        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(usuarioId))
        {
            return Challenge();
        }

        var membresias = await spService.PanelListarNegociosUsuarioAsync(usuarioId);

        if (membresias.Count == 0)
        {
            return View(new PanelDashboardViewModel
            {
                Mensaje = "Tu usuario aun no esta vinculado a un negocio."
            });
        }

        var negocioSeleccionadoId = negocioId ?? membresias[0].NegocioId;
        ViewData["AdminNegocioId"] = negocioSeleccionadoId;

        var contextoDashboard = await moduloPermisoService.ObtenerContextoAsync(User, negocioSeleccionadoId, "DASHBOARD");
        if (!contextoDashboard.Autorizado)
        {
            return View(new PanelDashboardViewModel
            {
                Mensaje = contextoDashboard.Mensaje,
                NegocioSeleccionadoId = negocioSeleccionadoId,
                Negocios = membresias
            });
        }

        var rolActual = await spService.PanelObtenerRolAsync(usuarioId, negocioSeleccionadoId);
        if (string.IsNullOrWhiteSpace(rolActual))
        {
            return RedirectToAction(nameof(Index), new { negocioId = membresias[0].NegocioId });
        }

        var sedes = await spService.EspaciosComboSedesAsync(negocioSeleccionadoId, contextoDashboard.EsAdministrador ? null : contextoDashboard.SedeIdAsignada);
        var sedeAplicada = ResolverSedeAplicada(contextoDashboard.EsAdministrador, contextoDashboard.SedeIdAsignada, sedeId, sedes);

        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today);
        var desde = fechaDesde ?? hasta.AddDays(-6);
        if (desde > hasta)
        {
            (desde, hasta) = (hasta, desde);
        }

        var diasPeriodo = (hasta.DayNumber - desde.DayNumber) + 1;
        var desdeAnterior = desde.AddDays(-diasPeriodo);
        var hastaAnterior = desde.AddDays(-1);

        var metricas = await spService.PanelObtenerMetricasAsync(negocioSeleccionadoId, hasta, sedeAplicada);
        var metricasDiaAnterior = await spService.PanelObtenerMetricasAsync(negocioSeleccionadoId, hasta.AddDays(-1), sedeAplicada);

        var ingresosPeriodo = await spService.ReportesIngresosPorDiaAsync(negocioSeleccionadoId, desde, hasta, sedeAplicada);
        var ingresosPeriodoAnterior = await spService.ReportesIngresosPorDiaAsync(negocioSeleccionadoId, desdeAnterior, hastaAnterior, sedeAplicada);
        var ocupacionPeriodo = await spService.ReportesOcupacionPorEspacioAsync(negocioSeleccionadoId, desde, hasta, sedeAplicada);

        var reservasPendientes = await spService.ReservasListarAsync(
            negocioSeleccionadoId,
            DateOnly.FromDateTime(DateTime.Today),
            DateOnly.FromDateTime(DateTime.Today.AddDays(2)),
            sedeAplicada,
            null,
            null,
            $"{(int)EstadoReserva.Pendiente},{(int)EstadoReserva.Confirmada}",
            1,
            20);

        var reservasSaldoPendiente = await spService.ReservasListarAsync(
            negocioSeleccionadoId,
            desde,
            hasta,
            sedeAplicada,
            null,
            null,
            $"{(int)EstadoReserva.Pendiente},{(int)EstadoReserva.Confirmada},{(int)EstadoReserva.Pagada}",
            1,
            200);

        var permisosRol = await spService.PanelListarModulosPermitidosAsync(usuarioId, negocioSeleccionadoId);
        var ingresosActualTotal = ingresosPeriodo.Sum(x => x.Ingresos);
        var ingresosAnteriorTotal = ingresosPeriodoAnterior.Sum(x => x.Ingresos);
        var reservasActualTotal = ingresosPeriodo.Sum(x => x.CantidadReservas);
        var reservasAnteriorTotal = ingresosPeriodoAnterior.Sum(x => x.CantidadReservas);

        var vm = new PanelDashboardViewModel
        {
            NegocioSeleccionadoId = negocioSeleccionadoId,
            Negocios = membresias,
            RolActual = rolActual,
            SedeId = sedeAplicada,
            SedesFiltro = PrepararSedesFiltro(sedes, contextoDashboard.EsAdministrador, sedeAplicada),
            FechaDesde = desde,
            FechaHasta = hasta,
            Modulos = permisosRol.Where(p => p.PuedeVer).OrderBy(p => p.ModuloNombre).ToList(),

            TotalSedes = metricas.TotalSedes,
            TotalEspacios = metricas.TotalEspacios,
            ReservasHoy = metricas.ReservasHoy,
            IngresosHoy = metricas.IngresosHoy,
            OcupacionHoyPct = metricas.OcupacionHoyPct,
            OcupacionDiaAnteriorPct = metricasDiaAnterior.OcupacionHoyPct,
            NoShowMes = metricas.NoShowMes,
            TicketPromedioMes = metricas.TicketPromedioMes,

            IngresosPeriodo = ingresosActualTotal,
            IngresosPeriodoAnterior = ingresosAnteriorTotal,
            ReservasPeriodo = reservasActualTotal,
            ReservasPeriodoAnterior = reservasAnteriorTotal,
            TicketPromedioPeriodo = reservasActualTotal > 0 ? ingresosActualTotal / reservasActualTotal : 0m,
            TicketPromedioPeriodoAnterior = reservasAnteriorTotal > 0 ? ingresosAnteriorTotal / reservasAnteriorTotal : 0m,

            SerieIngresosPorDia = ingresosPeriodo
                .OrderBy(x => x.Fecha)
                .Select(x => new DashboardSerieItemViewModel { Fecha = x.Fecha, Valor = x.Ingresos })
                .ToList(),

            SerieReservasPorDia = ingresosPeriodo
                .OrderBy(x => x.Fecha)
                .Select(x => new DashboardSerieItemViewModel { Fecha = x.Fecha, Valor = x.CantidadReservas })
                .ToList(),

            TopEspacios = ocupacionPeriodo
                .OrderByDescending(x => x.MontoCobrado)
                .ThenByDescending(x => x.CantidadReservas)
                .Take(6)
                .Select(x => new DashboardTopEspacioViewModel
                {
                    Sede = x.Sede,
                    Espacio = x.Espacio,
                    Reservas = x.CantidadReservas,
                    Horas = x.HorasReservadas,
                    Cobrado = x.MontoCobrado
                })
                .ToList(),

            ReservasPendientesConfirmacion = reservasPendientes.Reservas
                .OrderBy(x => x.Fecha)
                .ThenBy(x => x.HoraInicio)
                .Take(8)
                .Select(MapearReservaAccion)
                .ToList(),

            ReservasConSaldoPendiente = reservasSaldoPendiente.Reservas
                .Where(x => x.SaldoPendiente > 0)
                .OrderByDescending(x => x.SaldoPendiente)
                .ThenBy(x => x.Fecha)
                .Select(MapearReservaAccion)
                .ToList()
        };

        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerEspaciosReserva(int negocioId, int? sedeId)
    {
        var contextoDashboard = await moduloPermisoService.ObtenerContextoAsync(User, negocioId, "DASHBOARD");
        if (!contextoDashboard.Autorizado)
        {
            return Json(new { ok = false, mensaje = contextoDashboard.Mensaje ?? "No autorizado." });
        }

        var sedeAplicada = contextoDashboard.EsAdministrador ? sedeId : contextoDashboard.SedeIdAsignada;
        var espacios = await spService.ReservasComboEspaciosAsync(negocioId, sedeAplicada);
        return Json(new
        {
            ok = true,
            items = espacios.Select(x => new { value = x.Value, text = x.Text })
        });
    }

    private static int? ResolverSedeAplicada(bool esAdministrador, int? sedeAsignada, int? sedeSolicitada, IReadOnlyCollection<SelectListItem> sedesDisponibles)
    {
        if (!esAdministrador)
        {
            return sedeAsignada;
        }

        if (!sedeSolicitada.HasValue || sedeSolicitada.Value <= 0)
        {
            return null;
        }

        var sedeExiste = sedesDisponibles.Any(x => int.TryParse(x.Value, out var id) && id == sedeSolicitada.Value);
        return sedeExiste ? sedeSolicitada : null;
    }

    private static List<SelectListItem> PrepararSedesFiltro(List<SelectListItem> sedesBase, bool esAdministrador, int? sedeAplicada)
    {
        var result = new List<SelectListItem>();
        if (esAdministrador)
        {
            result.Add(new SelectListItem("Todas las sedes", string.Empty, !sedeAplicada.HasValue));
        }

        foreach (var sede in sedesBase)
        {
            if (!int.TryParse(sede.Value, out var sedeIdValor))
            {
                continue;
            }

            result.Add(new SelectListItem(sede.Text, sede.Value, sedeAplicada.HasValue && sedeAplicada.Value == sedeIdValor));
        }

        return result;
    }

    private static DashboardReservaAccionViewModel MapearReservaAccion(ReservaItemViewModel item)
    {
        return new DashboardReservaAccionViewModel
        {
            ReservaId = item.Id,
            ReservaCodigo = $"R-{item.Id:000000}",
            Cliente = item.Cliente,
            Sede = item.Sede,
            Espacio = item.Espacio,
            Fecha = item.Fecha,
            HoraInicio = item.HoraInicio,
            HoraFin = item.HoraFin,
            MontoTotal = item.Total,
            SaldoPendiente = item.SaldoPendiente,
            Estado = item.Estado
        };
    }
}
