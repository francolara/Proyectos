using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class SedesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new SedesIndexViewModel
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
            Sedes = await spService.SedesListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null))
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new SedeFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual, Activo = true };
        await CargarCatalogoServiciosAsync(vm);
        InicializarTelefonosParaVista(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(SedeFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        ComponerTelefonos(model);
        if (!ModelState.IsValid)
        {
            await CargarCatalogoServiciosAsync(model);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        await spService.SedesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.SedesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!SedePermitida(baseVm, vm.Id))
            return Forbid();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCatalogoServiciosAsync(vm);
        InicializarTelefonosParaVista(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(SedeFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!SedePermitida(baseVm, model.Id))
            return Forbid();
        ComponerTelefonos(model);
        if (!ModelState.IsValid)
        {
            await CargarCatalogoServiciosAsync(model);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        var ok = await spService.SedesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo actualizar la sede. Verifica el negocio seleccionado.");
            await CargarCatalogoServiciosAsync(model);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.SedesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private async Task CargarCatalogoServiciosAsync(SedeFormViewModel model)
    {
        model.ServiciosDisponibles = await spService.SedesComboServiciosAsync();
    }

    private static void InicializarTelefonosParaVista(SedeFormViewModel model)
    {
        TelefonoInternacionalHelper.Descomponer(model.Telefono, out var telefonoCodigoPais, out var telefonoNumeroLocal);
        model.TelefonoCodigoPais = telefonoCodigoPais;
        model.TelefonoNumeroLocal = telefonoNumeroLocal;

        TelefonoInternacionalHelper.Descomponer(model.WhatsappContacto, out var whatsappCodigoPais, out var whatsappNumeroLocal);
        model.WhatsappCodigoPais = whatsappCodigoPais;
        model.WhatsappNumeroLocal = whatsappNumeroLocal;
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
    }

    private static void ComponerTelefonos(SedeFormViewModel model)
    {
        model.Telefono = TelefonoInternacionalHelper.Componer(model.TelefonoCodigoPais, model.TelefonoNumeroLocal);
        model.WhatsappContacto = TelefonoInternacionalHelper.Componer(model.WhatsappCodigoPais, model.WhatsappNumeroLocal);
    }
}
