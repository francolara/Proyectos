using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ClientesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new ClientesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Clientes = await spService.ClientesListarAsync(resolvedNegocioId.Value)
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        return View(new ClienteFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            Activo = true,
            CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais("+51")
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        ComponerTelefono(model);
        if (!ModelState.IsValid)
        {
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        try
        {
            await spService.ClientesCrearAsync(model, User.Identity?.Name ?? "sistema");
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (SqlException ex) when (EsErrorClienteDuplicado(ex.Message))
        {
            ModelState.AddModelError(string.Empty, "Cliente ya se encuentra registrado.");
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ClientesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        InicializarTelefonoParaVista(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        ComponerTelefono(model);
        if (!ModelState.IsValid)
        {
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        try
        {
            var ok = await spService.ClientesActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar el cliente. Verifica el negocio seleccionado.");
                model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
                return View(model);
            }
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (SqlException ex) when (EsErrorClienteDuplicado(ex.Message))
        {
            ModelState.AddModelError(string.Empty, "Cliente ya se encuentra registrado.");
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.ClientesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private static void InicializarTelefonoParaVista(ClienteFormViewModel model)
    {
        TelefonoInternacionalHelper.Descomponer(model.Telefono, out var codigoPais, out var numeroLocal);
        model.TelefonoCodigoPais = codigoPais;
        model.TelefonoNumeroLocal = numeroLocal;
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
    }

    private static void ComponerTelefono(ClienteFormViewModel model)
    {
        model.Telefono = TelefonoInternacionalHelper.Componer(model.TelefonoCodigoPais, model.TelefonoNumeroLocal);
    }

    private static bool EsErrorClienteDuplicado(string? mensaje)
    {
        return !string.IsNullOrWhiteSpace(mensaje) &&
               mensaje.Contains("Cliente ya se encuentra registrado", StringComparison.OrdinalIgnoreCase);
    }
}
