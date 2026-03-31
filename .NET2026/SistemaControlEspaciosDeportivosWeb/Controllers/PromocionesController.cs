using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class PromocionesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new PromocionesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            SedeIdAsignada = baseVm.SedeIdAsignada,
            EsAdministrador = baseVm.EsAdministrador,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Promociones = await spService.PromocionesListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null))
        };

        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new PromocionFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            Activo = true
        };

        await CargarCombosAsync(vm, baseVm);
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            vm.SedeId = baseVm.SedeIdAsignada;
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(PromocionFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        await spService.PromocionesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.PromocionesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!SedePermitida(baseVm, vm.SedeId))
            return Forbid();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCombosAsync(vm, baseVm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(PromocionFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        var ok = await spService.PromocionesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar la promocion. Verifica el negocio seleccionado.");
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.PromocionesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private async Task CargarCombosAsync(PromocionFormViewModel model, ModuloBaseViewModel baseVm)
    {
        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId, sedeFiltro);
        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId, sedeFiltro);

        if (baseVm.EsAdministrador)
            model.Sedes.Insert(0, new SelectListItem("Todas las sedes", string.Empty));
        model.Espacios.Insert(0, new SelectListItem("Todos los espacios", string.Empty));
    }
}
