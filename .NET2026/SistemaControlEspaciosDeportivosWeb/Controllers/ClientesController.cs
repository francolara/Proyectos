using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.RegularExpressions;

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

        var vm = new ClienteFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            Activo = true,
            CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais("+51")
        };
        await CargarCombosClienteAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        ComponerTelefono(model);
        await NormalizarYValidarUbigeoAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCombosClienteAsync(model);
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
            await CargarCombosClienteAsync(model);
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
        await CargarCombosClienteAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        ComponerTelefono(model);
        await NormalizarYValidarUbigeoAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCombosClienteAsync(model);
            return View(model);
        }

        try
        {
            var ok = await spService.ClientesActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar el cliente. Verifica el negocio seleccionado.");
                await CargarCombosClienteAsync(model);
                return View(model);
            }
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (SqlException ex) when (EsErrorClienteDuplicado(ex.Message))
        {
            ModelState.AddModelError(string.Empty, "Cliente ya se encuentra registrado.");
            await CargarCombosClienteAsync(model);
            return View(model);
        }
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoProvincias(string? codigoDepartamento)
    {
        var codigoDep = (codigoDepartamento ?? string.Empty).Trim();
        if (codigoDep.Length != 2)
            return Json(Array.Empty<object>());

        var data = await spService.UbigeoProvinciasListarAsync(codigoDep);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoDistritos(string? codigoProvincia)
    {
        var codigoProv = (codigoProvincia ?? string.Empty).Trim();
        if (codigoProv.Length != 4)
            return Json(Array.Empty<object>());

        var data = await spService.UbigeoDistritosListarAsync(codigoProv);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
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

    private async Task CargarCombosClienteAsync(ClienteFormViewModel model)
    {
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
        model.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();

        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo) && Regex.IsMatch(model.CodigoUbigeo, @"^\d{6}$"))
        {
            model.CodigoDepartamento = model.CodigoUbigeo[..2];
            model.CodigoProvincia = model.CodigoUbigeo[..4];
        }

        model.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(model.CodigoDepartamento)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();

        model.DistritosUbigeo = !string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(model.CodigoProvincia)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
    }

    private async Task NormalizarYValidarUbigeoAsync(ClienteFormViewModel model)
    {
        model.DireccionFiscal = string.IsNullOrWhiteSpace(model.DireccionFiscal) ? null : model.DireccionFiscal.Trim();
        model.CodigoDepartamento = string.IsNullOrWhiteSpace(model.CodigoDepartamento) ? null : model.CodigoDepartamento.Trim();
        model.CodigoProvincia = string.IsNullOrWhiteSpace(model.CodigoProvincia) ? null : model.CodigoProvincia.Trim();
        model.CodigoUbigeo = string.IsNullOrWhiteSpace(model.CodigoUbigeo) ? null : model.CodigoUbigeo.Trim();

        if (string.IsNullOrWhiteSpace(model.DireccionFiscal))
        {
            model.CodigoDepartamento = null;
            model.CodigoProvincia = null;
            model.CodigoUbigeo = null;
            await CargarCombosClienteAsync(model);
            return;
        }

        if (string.IsNullOrWhiteSpace(model.CodigoDepartamento))
            ModelState.AddModelError(nameof(model.CodigoDepartamento), "Selecciona un departamento.");
        if (string.IsNullOrWhiteSpace(model.CodigoProvincia))
            ModelState.AddModelError(nameof(model.CodigoProvincia), "Selecciona una provincia.");
        if (string.IsNullOrWhiteSpace(model.CodigoUbigeo))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Selecciona un distrito.");

        if (!string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length != 2)
            ModelState.AddModelError(nameof(model.CodigoDepartamento), "Codigo de departamento invalido.");
        if (!string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length != 4)
            ModelState.AddModelError(nameof(model.CodigoProvincia), "Codigo de provincia invalido.");
        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo) && !Regex.IsMatch(model.CodigoUbigeo, @"^\d{6}$"))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Codigo de distrito invalido.");

        if (!string.IsNullOrWhiteSpace(model.CodigoDepartamento) &&
            !string.IsNullOrWhiteSpace(model.CodigoProvincia) &&
            !model.CodigoProvincia.StartsWith(model.CodigoDepartamento, StringComparison.Ordinal))
            ModelState.AddModelError(nameof(model.CodigoProvincia), "La provincia no corresponde al departamento seleccionado.");

        if (!string.IsNullOrWhiteSpace(model.CodigoProvincia) &&
            !string.IsNullOrWhiteSpace(model.CodigoUbigeo) &&
            !model.CodigoUbigeo.StartsWith(model.CodigoProvincia, StringComparison.Ordinal))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "El distrito no corresponde a la provincia seleccionada.");

        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo))
        {
            var ubigeo = await spService.UbigeoObtenerPorCodigoAsync(model.CodigoUbigeo);
            if (ubigeo is null)
            {
                ModelState.AddModelError(nameof(model.CodigoUbigeo), "El distrito seleccionado no existe.");
            }
            else
            {
                model.CodigoDepartamento = ubigeo.CodigoDepartamento;
                model.CodigoProvincia = ubigeo.CodigoProvincia;
            }
        }

        await CargarCombosClienteAsync(model);
    }
}
