using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize]
public class PanelController(ISportCenterStoredProcedureService spService, IModuloPermisoService moduloPermisoService) : Controller
{
    public async Task<IActionResult> Index(int? negocioId)
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

        var permisosRol = await spService.PanelListarModulosPermitidosAsync(usuarioId, negocioSeleccionadoId);

        var hoy = DateOnly.FromDateTime(DateTime.Today);
        var metricas = await spService.PanelObtenerMetricasAsync(negocioSeleccionadoId, hoy);

        var vm = new PanelDashboardViewModel
        {
            NegocioSeleccionadoId = negocioSeleccionadoId,
            Negocios = membresias,
            RolActual = rolActual,
            Modulos = permisosRol.Where(p => p.PuedeVer).OrderBy(p => p.ModuloNombre).ToList(),
            TotalSedes = metricas.TotalSedes,
            TotalEspacios = metricas.TotalEspacios,
            ReservasHoy = metricas.ReservasHoy,
            IngresosHoy = metricas.IngresosHoy,
            OcupacionHoyPct = metricas.OcupacionHoyPct,
            NoShowMes = metricas.NoShowMes,
            TicketPromedioMes = metricas.TicketPromedioMes
        };

        return View(vm);
    }
}
