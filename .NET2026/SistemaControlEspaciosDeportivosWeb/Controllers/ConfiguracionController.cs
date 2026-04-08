using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.RegularExpressions;

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
        vm.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        vm.Monedas = await spService.ConfiguracionClubComboMonedasAsync(negocioId);
        vm.PoliticasConfirmacionPago = ObtenerPoliticasConfirmacionPago();
        await CargarCombosUbigeoAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Index(ConfiguracionClubViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        model.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        model.Monedas = await spService.ConfiguracionClubComboMonedasAsync(model.NegocioId);
        model.PoliticasConfirmacionPago = ObtenerPoliticasConfirmacionPago();
        NormalizarYValidarPoliticaConfirmacionPago(model);
        await NormalizarYValidarUbigeoAsync(model);
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

    private async Task CargarCombosUbigeoAsync(ConfiguracionClubViewModel model)
    {
        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();

        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo) && Regex.IsMatch(model.CodigoUbigeo, @"^\d{6}$"))
        {
            model.CodigoDepartamento = model.CodigoUbigeo[..2];
            model.CodigoProvincia = model.CodigoUbigeo[..4];
        }

        model.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(model.CodigoDepartamento)
            : new List<SelectListItem>();

        model.DistritosUbigeo = !string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(model.CodigoProvincia)
            : new List<SelectListItem>();
    }

    private async Task NormalizarYValidarUbigeoAsync(ConfiguracionClubViewModel model)
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
            await CargarCombosUbigeoAsync(model);
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

        await CargarCombosUbigeoAsync(model);
    }

    private static List<SelectListItem> ObtenerPoliticasConfirmacionPago() =>
    [
        new SelectListItem("No exigir pago para confirmar", "0"),
        new SelectListItem("Exigir adelanto minimo (%)", "1"),
        new SelectListItem("Exigir pago total (100%)", "2")
    ];

    private void NormalizarYValidarPoliticaConfirmacionPago(ConfiguracionClubViewModel model)
    {
        if (model.PoliticaConfirmacionPago is < 0 or > 2)
        {
            ModelState.AddModelError(nameof(model.PoliticaConfirmacionPago), "La politica de confirmacion no es valida.");
            return;
        }

        if (model.PoliticaConfirmacionPago != 1)
        {
            model.PorcentajeAdelantoMinimo = null;
            return;
        }

        if (!model.PorcentajeAdelantoMinimo.HasValue)
        {
            ModelState.AddModelError(nameof(model.PorcentajeAdelantoMinimo), "Ingresa el porcentaje minimo de adelanto.");
            return;
        }

        if (model.PorcentajeAdelantoMinimo.Value < 1 || model.PorcentajeAdelantoMinimo.Value > 100)
        {
            ModelState.AddModelError(nameof(model.PorcentajeAdelantoMinimo), "El porcentaje minimo debe ser un numero entero entre 1 y 100.");
            return;
        }

        if (model.PorcentajeAdelantoMinimo.Value != Math.Truncate(model.PorcentajeAdelantoMinimo.Value))
        {
            ModelState.AddModelError(nameof(model.PorcentajeAdelantoMinimo), "El porcentaje minimo no admite decimales.");
        }
    }
}
