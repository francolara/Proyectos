using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class NotificacionesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    [HttpGet]
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public async Task<IActionResult> Resumen(int negocioId, int top = 15)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "DASHBOARD");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return Forbid();

        var safeTop = top <= 0 ? 15 : Math.Min(top, 50);
        var noLeidas = await spService.NotificacionesContarNoLeidasAsync(negocioId);
        var items = await spService.NotificacionesListarAsync(negocioId, safeTop);

        return Json(new
        {
            ok = true,
            noLeidas,
            items = items.Select(x => new
            {
                id = x.Id,
                tipo = x.Tipo,
                titulo = x.Titulo,
                mensaje = x.Mensaje,
                urlDestino = x.UrlDestino,
                fechaRegistroUtc = x.FechaRegistroUtc,
                fechaTexto = x.FechaRegistroUtc.ToLocalTime().ToString("dd/MM HH:mm")
            })
        });
    }
}
