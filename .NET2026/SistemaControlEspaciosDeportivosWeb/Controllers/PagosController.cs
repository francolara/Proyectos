using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class PagosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PAGOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new PagosIndexViewModel
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
            Pagos = await spService.PagosListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null))
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PAGOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new PagoFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        vm.Reservas = await spService.PagosComboReservasAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null));
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(PagoFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Reservas = await spService.PagosComboReservasAsync(model.NegocioId, AplicarSedeAsignada(baseVm, null));
        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        if (!ModelState.IsValid) return View(model);

        try
        {
            await spService.PagosCrearAsync(model, User.Identity?.Name ?? "sistema");
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PAGOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.PagosObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.Reservas = await spService.PagosComboReservasAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null));
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(PagoFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Reservas = await spService.PagosComboReservasAsync(model.NegocioId, AplicarSedeAsignada(baseVm, null));
        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        if (!ModelState.IsValid) return View(model);

        try
        {
            var ok = await spService.PagosActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar el pago. Verifica el negocio seleccionado.");
                return View(model);
            }
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.PagosEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }
}
