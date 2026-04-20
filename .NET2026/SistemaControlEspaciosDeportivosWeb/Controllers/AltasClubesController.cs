using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class AltasClubesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SOLICITUDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });
        var resultado = await spService.AltasClubesListarAsync(estado);

        var vm = new AltasClubesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = "Altas de clubes",
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Estado = estado,
            Solicitudes = resultado.Solicitudes
        };

        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Aprobar(int negocioId, AltaClubGestionFormViewModel model, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SOLICITUDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.AltasClubesAprobarAsync(model.Id, User.Identity?.Name ?? "sistema", model.ComentarioGestion, 30);
            TempData["MensajeAltaClub"] = "Solicitud aprobada y negocio/sede creados correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MensajeAltaClubError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId, estado });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Rechazar(int negocioId, AltaClubGestionFormViewModel model, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SOLICITUDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.AltasClubesRechazarAsync(model.Id, User.Identity?.Name ?? "sistema", model.ComentarioGestion);
            TempData["MensajeAltaClub"] = "Solicitud rechazada.";
        }
        catch (Exception ex)
        {
            TempData["MensajeAltaClubError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId, estado });
    }
}
