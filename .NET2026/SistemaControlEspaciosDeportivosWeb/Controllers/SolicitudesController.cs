using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class SolicitudesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SOLICITUDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today.AddDays(-7));
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today);
        if (hasta < desde) hasta = desde;

        var vm = new SolicitudesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            FechaDesde = desde,
            FechaHasta = hasta,
            Estado = estado,
            Solicitudes = await spService.SolicitudesPublicasListarAsync(negocioId, desde, hasta, estado)
        };

        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Aprobar(SolicitudEstadoFormViewModel model)
    {
        model.Estado = 2;
        return await ActualizarEstadoInterno(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Rechazar(SolicitudEstadoFormViewModel model)
    {
        model.Estado = 3;
        return await ActualizarEstadoInterno(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Convertir(SolicitudConvertirFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SOLICITUDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var reservaId = await spService.SolicitudesPublicasConvertirAReservaAsync(model, User.Identity?.Name ?? "sistema");
            TempData["MensajeSolicitudInterna"] = $"Solicitud convertida correctamente a reserva #{reservaId}.";
        }
        catch (Exception ex)
        {
            TempData["MensajeSolicitudInternaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    private async Task<IActionResult> ActualizarEstadoInterno(SolicitudEstadoFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SOLICITUDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.SolicitudesPublicasActualizarEstadoAsync(model, User.Identity?.Name ?? "sistema");
            TempData["MensajeSolicitudInterna"] = "Estado actualizado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MensajeSolicitudInternaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }
}
