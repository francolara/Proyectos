using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class EspaciosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new EspaciosIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Espacios = await spService.EspaciosListarAsync(negocioId)
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int negocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new EspacioFormViewModel { NegocioId = negocioId, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        vm.Sedes = await spService.EspaciosComboSedesAsync(negocioId);
        vm.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync();
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId);
        model.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync();
        if (!ModelState.IsValid) return View(model);

        await spService.EspaciosCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.EspaciosObtenerAsync(negocioId, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.Sedes = await spService.EspaciosComboSedesAsync(negocioId);
        vm.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync();
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId);
        model.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync();
        if (!ModelState.IsValid) return View(model);

        var ok = await spService.EspaciosActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.EspaciosEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }
}