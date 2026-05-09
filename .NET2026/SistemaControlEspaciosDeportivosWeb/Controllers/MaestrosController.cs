using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class MaestrosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "MAESTROS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new MaestrosIndexViewModel
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
            Monedas = await spService.MaestrosMonedasListarAsync(baseVm.NegocioId),
            MonedasSuper = await spService.MaestrosMonedasSuperListarAsync(),
            TiposSueloSuper = await spService.MaestrosTiposSueloSuperListarAsync(),
            TiposDeporteSuper = await spService.MaestrosTiposDeporteSuperListarAsync(),
            TiposSuelo = await spService.MaestrosTiposSueloListarAsync(baseVm.NegocioId),
            TiposDeporte = await spService.MaestrosTiposDeporteListarAsync(baseVm.NegocioId),
            FormasPago = await spService.MaestrosFormasPagoListarAsync(baseVm.NegocioId),
            TiposDocumentoComprobanteSuper = await spService.MaestrosTiposDocumentoComprobanteSuperListarAsync(),
            TiposDocumentoComprobante = await spService.MaestrosTiposDocumentoComprobanteListarAsync(baseVm.NegocioId)
        };
        var configClub = await spService.ConfiguracionClubObtenerAsync(baseVm.NegocioId);
        vm.EmisionComprobantesElectronicos = configClub?.EmisionComprobantesElectronicos ?? false;
        vm.EnviarComprobanteAutomatico = configClub?.EnviarComprobanteAutomatico ?? false;
        vm.EmisionReciboInterno = configClub?.EmisionReciboInterno ?? false;
        vm.TiposDocumentoComprobanteTributarios = await spService.CombosDocumentosComprobanteNegocioAsync(baseVm.NegocioId, true);
        vm.TiposDocumentoComprobanteNoTributarios = await spService.CombosDocumentosComprobanteNegocioAsync(baseVm.NegocioId, false);
        vm.SeriesDocumentoComprobante = await spService.ConfiguracionSeriesDocumentoListarAsync(baseVm.NegocioId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> MonedaCrear(int negocioId, int monedaSuperId, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.MaestrosMonedasCrearAsync(negocioId, monedaSuperId, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosOk"] = "Moneda registrada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> MonedaEditar(int negocioId, int id, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosMonedasActualizarAsync(negocioId, id, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo actualizar la moneda.";
            TempData["MaestrosOk"] = ok ? "Moneda actualizada correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> MonedaEliminar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosMonedasEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar la moneda.";
            TempData["MaestrosOk"] = ok ? "Moneda inactivada correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoSueloCrear(int negocioId, int tipoSueloSuperId, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.MaestrosTiposSueloCrearAsync(negocioId, tipoSueloSuperId, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosOk"] = "Tipo de suelo registrado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoSueloEditar(int negocioId, int id, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposSueloActualizarAsync(negocioId, id, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo actualizar el tipo de suelo.";
            TempData["MaestrosOk"] = ok ? "Tipo de suelo actualizado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoSueloEliminar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposSueloEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar el tipo de suelo.";
            TempData["MaestrosOk"] = ok ? "Tipo de suelo inactivado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDeporteCrear(int negocioId, int tipoDeporteSuperId, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.MaestrosTiposDeporteCrearAsync(negocioId, tipoDeporteSuperId, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosOk"] = "Tipo de deporte registrado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDeporteEditar(int negocioId, int id, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposDeporteActualizarAsync(negocioId, id, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo actualizar el tipo de deporte.";
            TempData["MaestrosOk"] = ok ? "Tipo de deporte actualizado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDeporteEliminar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposDeporteEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar el tipo de deporte.";
            TempData["MaestrosOk"] = ok ? "Tipo de deporte inactivado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> FormaPagoCrear(int negocioId, string nombre, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.MaestrosFormasPagoCrearAsync(negocioId, nombre, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosOk"] = "Forma de pago registrada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> FormaPagoEditar(int negocioId, int id, string nombre, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosFormasPagoActualizarAsync(negocioId, id, nombre, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo actualizar la forma de pago.";
            TempData["MaestrosOk"] = ok ? "Forma de pago actualizada correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> FormaPagoEliminar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosFormasPagoEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar la forma de pago.";
            TempData["MaestrosOk"] = ok ? "Forma de pago inactivada correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDocumentoComprobanteCrear(int negocioId, string codigoSunat, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            await spService.MaestrosTiposDocumentoComprobanteCrearAsync(negocioId, codigoSunat, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosOk"] = "Tipo de documento registrado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDocumentoComprobanteEditar(int negocioId, int id, bool activo = true)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposDocumentoComprobanteActualizarAsync(negocioId, id, activo, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo actualizar el tipo de documento.";
            TempData["MaestrosOk"] = ok ? "Tipo de documento actualizado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> TipoDocumentoComprobanteEliminar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.MaestrosTiposDocumentoComprobanteEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar el tipo de documento.";
            TempData["MaestrosOk"] = ok ? "Tipo de documento inactivado correctamente." : null;
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
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
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
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
            TempData["MaestrosOk"] = "Serie configurada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
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
        var baseVm = await ObtenerBaseAsync(negocioId, "MAESTROS");
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
            TempData["MaestrosOk"] = ok ? "Serie inactivada correctamente." : null;
            TempData["MaestrosError"] = ok ? null : "No se pudo inactivar la serie.";
        }
        catch (Exception ex)
        {
            TempData["MaestrosError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }
}
