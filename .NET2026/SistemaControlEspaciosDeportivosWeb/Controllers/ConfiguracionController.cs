using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.RegularExpressions;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ConfiguracionController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    ISedeImagenStorageService sedeImagenStorageService) : ModuloControllerBase(moduloPermisoService)
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
        await CargarConfigDocumentosSeriesAsync(vm);
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
        var configActual = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        var logoUrlActual = string.IsNullOrWhiteSpace(configActual?.LogoUrl) ? null : configActual!.LogoUrl!.Trim();
        string? logoUrlNuevo = null;

        model.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        model.Monedas = await spService.ConfiguracionClubComboMonedasAsync(model.NegocioId);
        model.PoliticasConfirmacionPago = ObtenerPoliticasConfirmacionPago();
        await CargarConfigDocumentosSeriesAsync(model);
        model.LogoUrl = logoUrlActual;

        if (model.QuitarLogo)
            model.LogoUrl = null;

        if (model.LogoArchivo is not null && model.LogoArchivo.Length > 0)
        {
            try
            {
                logoUrlNuevo = await sedeImagenStorageService.UploadLogoNegocioAsync(model.NegocioId, model.LogoArchivo, HttpContext.RequestAborted);
                if (string.IsNullOrWhiteSpace(logoUrlNuevo))
                {
                    ModelState.AddModelError(nameof(model.LogoArchivo), "No se pudo completar la carga del logo.");
                }
                else
                {
                    model.LogoUrl = logoUrlNuevo;
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(nameof(model.LogoArchivo), $"No se pudo subir el logo: {ex.Message}");
            }
        }

        NormalizarYValidarPoliticaConfirmacionPago(model);
        NormalizarYValidarCancelacionNoConfirmada(model);
        NormalizarYValidarIgv(model);
        await NormalizarYValidarUbigeoAsync(model);
        ValidarEmisionComprobantes(model);
        if (!ModelState.IsValid)
        {
            if (!string.IsNullOrWhiteSpace(logoUrlNuevo))
            {
                await sedeImagenStorageService.DeleteSedeImagenesAsync([logoUrlNuevo], HttpContext.RequestAborted);
                model.LogoUrl = logoUrlActual;
            }

            model.NegocioNombre = baseVm.NegocioNombre;
            model.RolActual = baseVm.RolActual;
            return View(model);
        }

        var ok = await spService.ConfiguracionClubActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            if (!string.IsNullOrWhiteSpace(logoUrlNuevo))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([logoUrlNuevo], HttpContext.RequestAborted);

            ModelState.AddModelError(string.Empty, "No se pudo actualizar la configuracion del club.");
            model.NegocioNombre = baseVm.NegocioNombre;
            model.RolActual = baseVm.RolActual;
            return View(model);
        }

        var debeEliminarLogoPrevio = !string.IsNullOrWhiteSpace(logoUrlActual) &&
            (model.QuitarLogo || (!string.IsNullOrWhiteSpace(logoUrlNuevo) &&
                                  !string.Equals(logoUrlActual, logoUrlNuevo, StringComparison.OrdinalIgnoreCase)));
        if (debeEliminarLogoPrevio)
            await sedeImagenStorageService.DeleteSedeImagenesAsync([logoUrlActual!], HttpContext.RequestAborted);

        TempData["ConfiguracionOk"] = "Configuracion del club actualizada correctamente.";
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarSerieDocumento(
        int negocioId,
        string codigoSunat,
        string serie,
        bool activo = true,
        bool emisionComprobantesElectronicos = false,
        bool enviarComprobanteAutomatico = false,
        bool emisionReciboInterno = false)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar)
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.ConfiguracionSeriesDocumentoGuardarAsync(negocioId, codigoSunat, serie, activo, User.Identity?.Name ?? "sistema");
            await spService.ConfiguracionClubActualizarEmisionAsync(
                negocioId,
                emisionComprobantesElectronicos,
                enviarComprobanteAutomatico,
                emisionReciboInterno,
                User.Identity?.Name ?? "sistema");
            TempData["ConfiguracionOk"] = "Serie configurada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["ConfiguracionError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarSerieDocumento(
        int negocioId,
        int id,
        bool emisionComprobantesElectronicos = false,
        bool enviarComprobanteAutomatico = false,
        bool emisionReciboInterno = false)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar)
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.ConfiguracionSeriesDocumentoEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            await spService.ConfiguracionClubActualizarEmisionAsync(
                negocioId,
                emisionComprobantesElectronicos,
                enviarComprobanteAutomatico,
                emisionReciboInterno,
                User.Identity?.Name ?? "sistema");
            TempData["ConfiguracionOk"] = ok ? "Serie inactivada correctamente." : null;
            TempData["ConfiguracionError"] = ok ? null : "No se pudo inactivar la serie.";
        }
        catch (Exception ex)
        {
            TempData["ConfiguracionError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
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

    private async Task CargarConfigDocumentosSeriesAsync(ConfiguracionClubViewModel model)
    {
        model.TiposDocumentoComprobanteTributarios = await spService.CombosDocumentosComprobanteNegocioAsync(model.NegocioId, true);
        model.TiposDocumentoComprobanteNoTributarios = await spService.CombosDocumentosComprobanteNegocioAsync(model.NegocioId, false);
        model.SeriesDocumentoComprobante = await spService.ConfiguracionSeriesDocumentoListarAsync(model.NegocioId);
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

    private void NormalizarYValidarIgv(ConfiguracionClubViewModel model)
    {
        if (model.PorcentajeIgv is < 0 or > 100)
            ModelState.AddModelError(nameof(model.PorcentajeIgv), "El porcentaje de IGV debe estar entre 0 y 100.");
    }

    private void NormalizarYValidarCancelacionNoConfirmada(ConfiguracionClubViewModel model)
    {
        if (!model.CancelacionAutomaticaNoConfirmada)
        {
            model.MinutosCancelacionNoConfirmada = null;
            return;
        }

        if (!model.MinutosCancelacionNoConfirmada.HasValue)
        {
            ModelState.AddModelError(nameof(model.MinutosCancelacionNoConfirmada),
                "Ingresa el tiempo de cancelacion automatica por no confirmacion.");
            return;
        }

        if (model.MinutosCancelacionNoConfirmada.Value <= 0)
        {
            ModelState.AddModelError(nameof(model.MinutosCancelacionNoConfirmada),
                "El tiempo de cancelacion automatica debe ser mayor a 0 minutos.");
        }
    }

    private void ValidarEmisionComprobantes(ConfiguracionClubViewModel model)
    {
        if (model.EmisionComprobantesElectronicos)
        {
            if (!string.Equals((model.TipoDocumento ?? string.Empty).Trim(), "6", StringComparison.Ordinal))
            {
                ModelState.AddModelError(nameof(model.TipoDocumento),
                    "Para activar emision de comprobantes electronicos, el tipo de documento del negocio debe ser RUC (6).");
            }

            if (model.PorcentajeIgv <= 0)
            {
                ModelState.AddModelError(nameof(model.PorcentajeIgv),
                    "Para activar emision de comprobantes electronicos, el IGV debe ser mayor que 0.");
            }

            if (string.IsNullOrWhiteSpace(model.DireccionFiscal))
            {
                ModelState.AddModelError(nameof(model.DireccionFiscal),
                    "Para activar emision de comprobantes electronicos, la direccion fiscal es obligatoria.");
            }

            if (string.IsNullOrWhiteSpace(model.CodigoUbigeo))
            {
                ModelState.AddModelError(nameof(model.CodigoUbigeo),
                    "Para activar emision de comprobantes electronicos, el ubigeo fiscal es obligatorio.");
            }
        }

        // La configuración de series por documento se valida y administra desde Maestros.
    }
}
