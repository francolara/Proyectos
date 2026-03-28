using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ComprobantesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new ComprobantesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Comprobantes = await spService.ComprobantesListarAsync(resolvedNegocioId.Value)
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new ComprobanteFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        vm.Reservas = await spService.ComprobantesComboReservasAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ComprobanteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Reservas = await spService.ComprobantesComboReservasAsync(model.NegocioId);
        if (!ModelState.IsValid) return View(model);

        await spService.ComprobantesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ComprobantesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.Reservas = await spService.ComprobantesComboReservasAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ComprobanteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Reservas = await spService.ComprobantesComboReservasAsync(model.NegocioId);
        if (!ModelState.IsValid) return View(model);

        var ok = await spService.ComprobantesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar el comprobante. Verifica el negocio seleccionado.");
            return View(model);
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.ComprobantesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }
}
