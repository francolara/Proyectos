using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ConfiguracionController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService) : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SEDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = await spService.ConfiguracionClubObtenerAsync(negocioId) ?? new ConfiguracionClubViewModel { NegocioId = negocioId, Id = negocioId };
        vm.NegocioId = baseVm.NegocioId;
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.ModuloCodigo = baseVm.ModuloCodigo;
        vm.ModuloNombre = "Configuracion";
        vm.PuedeCrear = baseVm.PuedeCrear;
        vm.PuedeEditar = baseVm.PuedeEditar;
        vm.PuedeEliminar = baseVm.PuedeEliminar;
        vm.TiposDocumento = ObtenerTiposDocumento();
        vm.Monedas = await spService.ConfiguracionClubComboMonedasAsync(negocioId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Index(ConfiguracionClubViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        model.TiposDocumento = ObtenerTiposDocumento();
        model.Monedas = await spService.ConfiguracionClubComboMonedasAsync(model.NegocioId);
        if (!ModelState.IsValid)
        {
            model.NegocioNombre = baseVm.NegocioNombre;
            model.RolActual = baseVm.RolActual;
            return View(model);
        }

        var ok = await spService.ConfiguracionClubActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo actualizar la configuracion del club.");
            model.NegocioNombre = baseVm.NegocioNombre;
            model.RolActual = baseVm.RolActual;
            return View(model);
        }

        TempData["ConfiguracionOk"] = "Configuracion del club actualizada correctamente.";
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    private static List<SelectListItem> ObtenerTiposDocumento()
        =>
        [
            new("DNI", "DNI"),
            new("RUC", "RUC"),
            new("CE", "CE"),
            new("Pasaporte", "PASAPORTE")
        ];
}
