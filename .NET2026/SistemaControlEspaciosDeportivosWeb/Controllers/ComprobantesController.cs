using System.Text.RegularExpressions;
using System.Text;
using System.IO.Compression;
using System.Data;
using System.Net;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Globalization;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ComprobantesController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    IComprobanteElectronicoEmisionService emisionService,
    IEmailService emailService,
    IConfiguration configuration,
    IHttpClientFactory httpClientFactory)
    : ModuloControllerBase(moduloPermisoService)
{
    private readonly string _connectionString = configuration.GetConnectionString("DefaultConnection")
                                               ?? throw new InvalidOperationException("No se encontro DefaultConnection.");
    public async Task<IActionResult> Index(int? negocioId, string? buscar = null, string? codigoDocumento = null, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, string? preset = null, int pagina = 1)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var codigoDocumentoFiltro = string.IsNullOrWhiteSpace(codigoDocumento) ? null : codigoDocumento.Trim().ToUpperInvariant();
        var tiposDocumentoFiltro = await spService.CombosDocumentosComprobanteNegocioAsync(resolvedNegocioId.Value);
        if (!string.IsNullOrWhiteSpace(codigoDocumentoFiltro) && !tiposDocumentoFiltro.Any(x => string.Equals(x.Value, codigoDocumentoFiltro, StringComparison.OrdinalIgnoreCase)))
        {
            codigoDocumentoFiltro = null;
        }

        var (desde, hasta) = ResolverRangoFechas(fechaDesde, fechaHasta, preset);
        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var (comprobantes, totalRegistros) = await spService.ComprobantesListarAsync(
            resolvedNegocioId.Value,
            AplicarSedeAsignada(baseVm, null),
            buscar,
            codigoDocumentoFiltro,
            desde,
            hasta,
            paginaActual,
            tamanoPagina);
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
        {
            paginaActual = totalPaginas;
            (comprobantes, totalRegistros) = await spService.ComprobantesListarAsync(
                resolvedNegocioId.Value,
                AplicarSedeAsignada(baseVm, null),
                buscar,
                codigoDocumentoFiltro,
                desde,
                hasta,
                paginaActual,
                tamanoPagina);
        }
        var (totalMontoEmitidoGeneral, totalPendientesGeneral, totalAnuladosGeneral) = await CalcularTotalesComprobantesAsync(
            resolvedNegocioId.Value,
            AplicarSedeAsignada(baseVm, null),
            buscar,
            codigoDocumentoFiltro,
            desde,
            hasta);

        var vm = new ComprobantesIndexViewModel
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
            Buscar = buscar,
            CodigoDocumento = codigoDocumentoFiltro,
            FechaDesde = desde,
            FechaHasta = hasta,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            TotalMontoEmitidoGeneral = totalMontoEmitidoGeneral,
            TotalPendientesGeneral = totalPendientesGeneral,
            TotalAnuladosGeneral = totalAnuladosGeneral,
            TiposDocumentoFiltro = tiposDocumentoFiltro,
            Comprobantes = comprobantes
        };
        return View(vm);
    }

    private async Task<(decimal TotalMonto, int TotalPendientes, int TotalAnulados)> CalcularTotalesComprobantesAsync(
        int negocioId,
        int? sedeId,
        string? buscar,
        string? codigoDocumento,
        DateOnly desde,
        DateOnly hasta)
    {
        const int tamanoLote = 500;
        var pagina = 1;
        var totalMonto = 0m;
        var totalPendientes = 0;
        var totalAnulados = 0;
        var totalRegistros = 0;

        do
        {
            var (items, total) = await spService.ComprobantesListarAsync(
                negocioId,
                sedeId,
                buscar,
                codigoDocumento,
                desde,
                hasta,
                pagina,
                tamanoLote);
            totalRegistros = total;

            foreach (var item in items)
            {
                totalMonto += item.Total;
                if (item.Estado.Contains("Pend", StringComparison.OrdinalIgnoreCase))
                    totalPendientes++;
                if (item.Estado.Contains("Anul", StringComparison.OrdinalIgnoreCase))
                    totalAnulados++;
            }

            if (items.Count == 0)
                break;

            pagina++;
        } while ((pagina - 1) * tamanoLote < totalRegistros);

        return (totalMonto, totalPendientes, totalAnulados);
    }

    private static (DateOnly Desde, DateOnly Hasta) ResolverRangoFechas(DateOnly? fechaDesde, DateOnly? fechaHasta, string? preset)
    {
        var hoy = DateOnly.FromDateTime(DateTime.Today);
        DateOnly desde;
        DateOnly hasta;

        switch ((preset ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "hoy":
                desde = hoy;
                hasta = hoy;
                break;
            case "7d":
                hasta = hoy;
                desde = hoy.AddDays(-6);
                break;
            case "30d":
                hasta = hoy;
                desde = hoy.AddDays(-29);
                break;
            case "mes":
                desde = new DateOnly(hoy.Year, hoy.Month, 1);
                hasta = new DateOnly(hoy.Year, hoy.Month, DateTime.DaysInMonth(hoy.Year, hoy.Month));
                break;
            default:
                desde = fechaDesde ?? hoy.AddDays(-6);
                hasta = fechaHasta ?? hoy;
                break;
        }

        if (hasta < desde)
            (desde, hasta) = (hasta, desde);

        return (desde, hasta);
    }

    public async Task<IActionResult> Create(int? negocioId, int? reservaId = null, string? codigoSunat = null)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var config = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var vm = new ComprobanteFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ReservaId = reservaId ?? 0,
            FechaEmision = DateTime.Today,
            EmisionComprobantesElectronicos = config?.EmisionComprobantesElectronicos == true,
            EmisionReciboInterno = config?.EmisionReciboInterno == true,
            ForzarFiltroDocumentoPorCodigo = reservaId.HasValue && reservaId.Value > 0 && !string.IsNullOrWhiteSpace(codigoSunat),
            PorcentajeIgvConfigurado = config?.PorcentajeIgv ?? 18,
            CodigoDocumentoComprobante = string.IsNullOrWhiteSpace(codigoSunat) ? "03" : codigoSunat.Trim().ToUpperInvariant()
        };
        await PoblarDatosComprobanteAsync(vm, baseVm, reservaId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ComprobanteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        await PoblarDatosComprobanteAsync(model, baseVm, model.ReservaId);
        AplicarCalculoComprobante(model);
        if (!model.DocumentoTributario)
        {
            model.Igv = 0m;
            model.SubTotal = model.Total;
        }
        if (!model.TiposDocumentoComprobante.Any())
        {
            ModelState.AddModelError(string.Empty, "No hay tipos de documento habilitados para emitir comprobantes.");
        }
        if (model.ReservaId <= 0)
        {
            ModelState.AddModelError(nameof(model.ReservaId), "Selecciona una reserva pagada valida.");
        }
        else if (!model.Reservas.Any(x => int.TryParse(x.Value, out var id) && id == model.ReservaId))
        {
            ModelState.AddModelError(nameof(model.ReservaId), "La reserva seleccionada ya no esta disponible para emitir comprobante.");
        }
        if (!model.NegocioSerieId.HasValue || model.NegocioSerieId <= 0)
        {
            ModelState.AddModelError(nameof(model.NegocioSerieId), "Selecciona una serie valida para el documento.");
        }
        NormalizarDatosCliente(model);
        var montoMaximoBoletaSinDoc = await ObtenerMontoMaximoBoletaSinDocAsync();
        ValidarReglasDocumento(model, montoMaximoBoletaSinDoc);
        model.TipoComprobante = MapearTipoComprobante(model.CodigoDocumentoComprobante);
        if (!ModelState.IsValid) return View(model);

        var usuarioActual = User.Identity?.Name ?? "sistema";
        int id;
        try
        {
            id = await spService.ComprobantesCrearAsync(model, usuarioActual);
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }

        var resultadoEmision = await emisionService.EmitirAsync(model.NegocioId, id, usuarioActual);
        var comprobanteCreado = await spService.ComprobantesObtenerAsync(model.NegocioId, id);
        if (comprobanteCreado is not null)
        {
            TempData["ComprobanteOk"] = $"Comprobante emitido: {comprobanteCreado.Serie}-{comprobanteCreado.Numero:D8}.";
            if (!resultadoEmision.Exito)
            {
                TempData["ComprobanteError"] = string.Equals(resultadoEmision.Codigo, "PROVEEDOR_NO_IMPLEMENTADO", StringComparison.OrdinalIgnoreCase)
                    ? "Proveedor SUNAT no configurado para envio automatico."
                    : $"Se genero el comprobante, pero fallo el envio al proveedor: {resultadoEmision.Mensaje}";
            }
        }
        return RedirectToAction(nameof(Preview), new { negocioId = model.NegocioId, id });
    }

    public async Task<IActionResult> CreateNota(int id, string tipoNota, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var referencia = await spService.ComprobantesObtenerAsync(resolvedNegocioId.Value, id);
        if (referencia is null) return NotFound();
        if (!EsComprobanteAptoParaNota(referencia))
        {
            TempData["ComprobanteError"] = "Solo se puede generar NC/ND desde comprobantes Factura o Boleta aceptados por SUNAT.";
            return RedirectToAction(nameof(Index), new { negocioId = resolvedNegocioId.Value });
        }

        var tipoNotaNormalizado = NormalizarTipoNota(tipoNota);
        if (tipoNotaNormalizado is null)
            return BadRequest("Tipo de nota no valido.");

        var codigoDocumentoNota = CodigoDocumentoPorTipoNota(tipoNotaNormalizado);
        var vm = new ComprobanteFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ReservaId = referencia.ReservaId,
            FechaEmision = DateTime.Today,
            CodigoDocumentoComprobante = codigoDocumentoNota,
            TipoComprobante = MapearTipoComprobante(codigoDocumentoNota),
            DocumentoTributario = true,
            Total = referencia.Total,
            SubTotal = referencia.SubTotal,
            Igv = referencia.Igv,
            ClienteCorreo = referencia.ClienteCorreo,
            ClienteTipoDocumento = referencia.ClienteTipoDocumento,
            ClienteNumeroDocumento = referencia.ClienteNumeroDocumento,
            ClienteDireccionFiscal = referencia.ClienteDireccionFiscal,
            ClienteCodigoUbigeo = referencia.ClienteCodigoUbigeo,
            TipoMoneda = referencia.TipoMoneda,
            EsNota = true,
            TipoNota = tipoNotaNormalizado,
            ComprobanteReferenciaId = referencia.Id,
            ComprobanteReferenciaTipo = NombreDocumentoPorCodigo(referencia.CodigoDocumentoComprobante),
            ComprobanteReferenciaSerie = referencia.Serie,
            ComprobanteReferenciaNumero = referencia.Numero
        };

        await PoblarDatosComprobanteAsync(vm, baseVm, vm.ReservaId);
        vm.TipoMoneda = referencia.TipoMoneda;
        vm.CodigoDocumentoComprobante = codigoDocumentoNota;
        vm.TipoComprobante = MapearTipoComprobante(codigoDocumentoNota);
        vm.TiposDocumentoComprobante = new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>
        {
            new(
                $"{NombreDocumentoPorCodigo(codigoDocumentoNota)} ({codigoDocumentoNota})",
                codigoDocumentoNota)
        };

        vm.TiposNotaSunat = await spService.CombosTiposNotaComprobanteSunatAsync(tipoNotaNormalizado);
        vm.TipoNotaCodigoSunat = vm.TiposNotaSunat.FirstOrDefault()?.Value;
        if (!vm.SeriesDocumento.Any())
        {
            ModelState.AddModelError(nameof(vm.NegocioSerieId), "No hay series habilitadas para este tipo de nota en la sede de la reserva.");
        }
        AplicarCalculoComprobante(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CreateNota(ComprobanteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.NegocioNombre = baseVm.NegocioNombre;
        model.RolActual = baseVm.RolActual;
        model.EsNota = true;

        var tipoNotaNormalizado = NormalizarTipoNota(model.TipoNota);
        if (tipoNotaNormalizado is null)
        {
            ModelState.AddModelError(nameof(model.TipoNota), "Tipo de nota no valido.");
            tipoNotaNormalizado = "NC";
        }

        model.TipoNota = tipoNotaNormalizado;
        var codigoDocumentoNota = CodigoDocumentoPorTipoNota(tipoNotaNormalizado);
        model.CodigoDocumentoComprobante = codigoDocumentoNota;
        model.TipoComprobante = MapearTipoComprobante(codigoDocumentoNota);
        model.DocumentoTributario = true;

        if (!model.ComprobanteReferenciaId.HasValue || model.ComprobanteReferenciaId <= 0)
        {
            ModelState.AddModelError(nameof(model.ComprobanteReferenciaId), "Selecciona un comprobante de referencia valido.");
        }

        ComprobanteFormViewModel? referencia = null;
        if (model.ComprobanteReferenciaId.HasValue && model.ComprobanteReferenciaId > 0)
        {
            referencia = await spService.ComprobantesObtenerAsync(model.NegocioId, model.ComprobanteReferenciaId.Value);
            if (referencia is null || !EsComprobanteAptoParaNota(referencia))
            {
                ModelState.AddModelError(nameof(model.ComprobanteReferenciaId), "El comprobante de referencia no es valido para emitir notas.");
            }
            else
            {
                model.ReservaId = referencia.ReservaId;
                model.ComprobanteReferenciaTipo = NombreDocumentoPorCodigo(referencia.CodigoDocumentoComprobante);
                model.ComprobanteReferenciaSerie = referencia.Serie;
                model.ComprobanteReferenciaNumero = referencia.Numero;
                if (model.Total <= 0)
                {
                    model.Total = referencia.Total;
                }
            }
        }

        await PoblarDatosComprobanteAsync(model, baseVm, model.ReservaId);
        if (referencia is not null)
        {
            model.TipoMoneda = referencia.TipoMoneda;
        }
        model.CodigoDocumentoComprobante = codigoDocumentoNota;
        model.TipoComprobante = MapearTipoComprobante(codigoDocumentoNota);
        model.TiposDocumentoComprobante = new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>
        {
            new(
                $"{NombreDocumentoPorCodigo(codigoDocumentoNota)} ({codigoDocumentoNota})",
                codigoDocumentoNota)
        };
        model.TiposNotaSunat = await spService.CombosTiposNotaComprobanteSunatAsync(tipoNotaNormalizado);

        if (!model.TiposNotaSunat.Any())
        {
            ModelState.AddModelError(string.Empty, "No existen tipos de nota configurados para SUNAT.");
        }
        else if (string.IsNullOrWhiteSpace(model.TipoNotaCodigoSunat) || !model.TiposNotaSunat.Any(x => string.Equals(x.Value, model.TipoNotaCodigoSunat, StringComparison.OrdinalIgnoreCase)))
        {
            ModelState.AddModelError(nameof(model.TipoNotaCodigoSunat), "Selecciona el tipo de nota SUNAT.");
        }

        if (!model.NegocioSerieId.HasValue || model.NegocioSerieId <= 0 || !model.SeriesDocumento.Any(x => x.Value == model.NegocioSerieId.Value.ToString(CultureInfo.InvariantCulture)))
        {
            ModelState.AddModelError(nameof(model.NegocioSerieId), "Selecciona una serie valida para la nota.");
        }
        else
        {
            var serieElegida = model.SeriesDocumento.FirstOrDefault(x => x.Value == model.NegocioSerieId.Value.ToString(CultureInfo.InvariantCulture));
            model.Serie = serieElegida?.Text ?? model.Serie;
        }

        ValidarSerieNotaSegunReferencia(model, referencia);

        AplicarCalculoComprobante(model);
        NormalizarDatosCliente(model);
        var montoMaximoBoletaSinDoc = await ObtenerMontoMaximoBoletaSinDocAsync();
        ValidarReglasDocumento(model, montoMaximoBoletaSinDoc);
        if (!ModelState.IsValid) return View(model);

        var usuarioActual = User.Identity?.Name ?? "sistema";
        int idGenerado;
        try
        {
            idGenerado = await spService.ComprobantesCrearAsync(model, usuarioActual);
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
        var resultadoEmision = await emisionService.EmitirAsync(model.NegocioId, idGenerado, usuarioActual);
        var comprobanteCreado = await spService.ComprobantesObtenerAsync(model.NegocioId, idGenerado);
        if (comprobanteCreado is not null)
        {
            TempData["ComprobanteOk"] = $"{(tipoNotaNormalizado == "NC" ? "Nota de credito" : "Nota de debito")} generada: {comprobanteCreado.Serie}-{comprobanteCreado.Numero:D8}.";
            if (!resultadoEmision.Exito)
            {
                TempData["ComprobanteError"] = string.Equals(resultadoEmision.Codigo, "PROVEEDOR_NO_IMPLEMENTADO", StringComparison.OrdinalIgnoreCase)
                    ? "Proveedor SUNAT no configurado para envio automatico."
                    : $"Se genero el comprobante, pero fallo el envio al proveedor: {resultadoEmision.Mensaje}";
            }
        }

        return RedirectToAction(nameof(Preview), new { negocioId = model.NegocioId, id = idGenerado });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ComprobantesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        vm.EsEdicion = true;
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await PoblarDatosComprobanteAsync(vm, baseVm, vm.ReservaId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ComprobanteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var actual = await spService.ComprobantesObtenerAsync(model.NegocioId, model.Id);
        if (actual is null) return NotFound();

        var esPendiente = actual.Estado == SistemaControlEspaciosDeportivosWeb.Models.EstadoComprobanteElectronico.PendienteEnvio;
        if (!esPendiente)
        {
            ModelState.AddModelError(string.Empty, "Solo se permite editar datos del cliente cuando el comprobante esta pendiente.");
            actual.NegocioNombre = baseVm.NegocioNombre;
            actual.RolActual = baseVm.RolActual;
            actual.EsEdicion = true;
            await PoblarDatosComprobanteAsync(actual, baseVm, actual.ReservaId);
            return View(actual);
        }

        model.ReservaId = actual.ReservaId;
        model.TipoComprobante = actual.TipoComprobante;
        model.CodigoDocumentoComprobante = actual.CodigoDocumentoComprobante;
        model.Serie = actual.Serie;
        model.Numero = actual.Numero;
        model.FechaEmision = actual.FechaEmision;
        model.TipoMoneda = actual.TipoMoneda;
        model.SubTotal = actual.SubTotal;
        model.Igv = actual.Igv;
        model.Total = actual.Total;
        model.Estado = actual.Estado;
        model.EsEdicion = true;

        await PoblarDatosComprobanteAsync(model, baseVm, actual.ReservaId);
        AplicarCalculoComprobante(model);
        model.TipoComprobante = MapearTipoComprobante(model.CodigoDocumentoComprobante);
        NormalizarDatosCliente(model);
        var montoMaximoBoletaSinDoc = await ObtenerMontoMaximoBoletaSinDocAsync();
        ValidarReglasDocumento(model, montoMaximoBoletaSinDoc);
        if (!ModelState.IsValid) return View(model);

        bool ok;
        try
        {
            ok = await spService.ComprobantesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar el comprobante. Verifica el negocio seleccionado.");
            return View(model);
        }
        TempData["ComprobanteOk"] = $"Comprobante actualizado: {model.Serie}-{model.Numero:D8}.";
        return RedirectToAction(nameof(Edit), new { negocioId = model.NegocioId, id = model.Id });
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

    [HttpGet]
    public async Task<IActionResult> BuscarReservas(int negocioId, string? q = null, int? reservaId = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var items = await spService.ComprobantesBuscarReservasPagadasAsync(negocioId, q, reservaId, 40);
        return Json(new
        {
            ok = true,
            items = items.Select(x => new { value = x.Value, text = x.Text })
        });
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerContextoReserva(int negocioId, int reservaId, string? codigoSunat = null, bool forzarFiltroCodigo = false)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var config = await spService.ConfiguracionClubObtenerAsync(negocioId);
        var data = await spService.ComprobantesObtenerContextoReservaAsync(negocioId, reservaId, codigoSunat);
        if (data is null) return Json(new { ok = false, mensaje = "No se encontro la reserva pagada seleccionada." });

        var codigo = (codigoSunat ?? string.Empty).Trim().ToUpperInvariant();
        var documentosFiltrados = data.DocumentosDisponibles;
        if (forzarFiltroCodigo && codigo is "01" or "03")
        {
            documentosFiltrados = data.DocumentosDisponibles
                .Where(x =>
                {
                    var v = (x.Value ?? string.Empty).Trim().ToUpperInvariant();
                    return v is "01" or "03";
                })
                .ToList();
        }
        else if (forzarFiltroCodigo && codigo == "RI")
        {
            documentosFiltrados = data.DocumentosDisponibles
                .Where(x => string.Equals((x.Value ?? string.Empty).Trim(), "RI", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
        else
        {
            documentosFiltrados = data.DocumentosDisponibles
                .Where(x =>
                {
                    var v = (x.Value ?? string.Empty).Trim().ToUpperInvariant();
                    return v is not ("07" or "08");
                })
                .ToList();
        }

        return Json(new
        {
            ok = true,
            data = new
            {
                reservaId = data.ReservaId,
                reservaCodigo = data.ReservaCodigo,
                sede = data.Sede,
                espacio = data.Espacio,
                cliente = data.Cliente,
                clienteId = data.ClienteId,
                clienteCorreo = data.ClienteCorreo,
                clienteTipoDocumento = data.ClienteTipoDocumento,
                clienteNumeroDocumento = data.ClienteNumeroDocumento,
                clienteDireccionFiscal = data.ClienteDireccionFiscal,
                clienteCodigoDepartamento = data.ClienteCodigoDepartamento,
                clienteCodigoProvincia = data.ClienteCodigoProvincia,
                clienteCodigoUbigeo = data.ClienteCodigoUbigeo,
                fechaReserva = data.FechaReserva.ToString("yyyy-MM-dd"),
                horaInicioReserva = data.HoraInicioReserva.ToString("HH\\:mm"),
                horaFinReserva = data.HoraFinReserva.ToString("HH\\:mm"),
                totalReserva = data.TotalReserva,
                totalPagado = data.TotalPagado,
                saldoPendiente = data.SaldoPendiente,
                monedaSimbolo = data.MonedaSimbolo,
                porcentajeIgv = config?.PorcentajeIgv ?? 18,
                documentos = documentosFiltrados.Select(x => new { value = x.Value, text = x.Text }).ToList(),
                series = data.SeriesDisponibles.Select(x => new { value = x.Value, text = x.Text }).ToList(),
                pagos = data.PagosReserva.Select(p => new
                {
                    pagoId = p.PagoId,
                    fechaPago = p.FechaPago.ToString("dd/MM/yyyy"),
                    monto = p.Monto,
                    formaPago = p.FormaPago,
                    numeroOperacion = p.NumeroOperacion,
                    observacion = p.Observacion
                }).ToList()
            }
        });
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

    [HttpGet]
    public async Task<IActionResult> Preview(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var data = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (data is null) return NotFound();
        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, id);
        ViewBag.EstadoComprobante = comprobanteBase?.Estado;

        return View(data);
    }

    [HttpGet]
    public async Task<IActionResult> ImprimirInterno(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var data = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (data is null) return NotFound();
        if (!string.Equals(data.CodigoDocumentoComprobante?.Trim(), "RI", StringComparison.OrdinalIgnoreCase))
        {
            TempData["ComprobanteError"] = "La impresion HTML interna solo esta disponible para recibos internos.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        return View(data);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EnviarSunat(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var usuarioActual = User.Identity?.Name ?? "sistema";
        var resultado = await emisionService.EmitirManualAsync(negocioId, id, usuarioActual);
        TempData[resultado.Exito ? "ComprobanteOk" : "ComprobanteError"] = resultado.Mensaje;
        return RedirectToAction(nameof(Preview), new { negocioId, id });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ImprimirTributario(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var usuarioActual = User.Identity?.Name ?? "sistema";
        var resultadoConsulta = await emisionService.ConsultarEstadoAsync(negocioId, id, usuarioActual);
        if (!resultadoConsulta.Exito)
        {
            TempData["ComprobanteError"] = $"No se pudo consultar estado en proveedor: {resultadoConsulta.Mensaje}";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        var data = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (data is null) return NotFound();

        if (!string.IsNullOrWhiteSpace(data.UrlDescargaProveedor))
        {
            return Redirect(data.UrlDescargaProveedor);
        }

        TempData["ComprobanteError"] = "El proveedor no devolvio URL de descarga para este comprobante.";
        return RedirectToAction(nameof(Preview), new { negocioId, id });
    }

    [HttpGet]
    public async Task<IActionResult> DescargarDocsSunat(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var comprobante = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (comprobante is null) return NotFound();

        var urls = await ObtenerUrlsDocumentosSunatAsync(negocioId, id);
        var pares = new List<(string Nombre, string Url)>();

        if (!string.IsNullOrWhiteSpace(urls.UrlPdfSunat)) pares.Add(("PDF", urls.UrlPdfSunat!));
        if (!string.IsNullOrWhiteSpace(urls.UrlXmlSunat)) pares.Add(("XML", urls.UrlXmlSunat!));
        if (!string.IsNullOrWhiteSpace(urls.UrlCdrSunat)) pares.Add(("CDR", urls.UrlCdrSunat!));

        if (pares.Count == 0)
        {
            TempData["ComprobanteError"] = "No hay URLs PDF/XML/CDR disponibles para este comprobante.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        var client = httpClientFactory.CreateClient();
        client.Timeout = TimeSpan.FromSeconds(30);

        await using var zipStream = new MemoryStream();
        using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (nombre, url) in pares)
            {
                if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
                    continue;

                byte[] contenido;
                try
                {
                    contenido = await client.GetByteArrayAsync(uri);
                }
                catch
                {
                    continue;
                }

                var extension = Path.GetExtension(uri.AbsolutePath);
                if (string.IsNullOrWhiteSpace(extension))
                {
                    extension = nombre switch
                    {
                        "PDF" => ".pdf",
                        "XML" => ".xml",
                        _ => ".zip"
                    };
                }

                var nombreBase = $"{comprobante.Serie}-{comprobante.Numero:D8}";
                var entry = zip.CreateEntry($"{nombreBase}_{nombre}{extension}", CompressionLevel.Optimal);
                await using var entryStream = entry.Open();
                await entryStream.WriteAsync(contenido);
            }
        }

        if (zipStream.Length == 0)
        {
            TempData["ComprobanteError"] = "No se pudo descargar ningun documento desde las URLs registradas.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        zipStream.Position = 0;
        return File(zipStream.ToArray(), "application/zip", $"{comprobante.Serie}-{comprobante.Numero:D8}_SUNAT.zip");
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EnviarComprobanteCorreo(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var comprobante = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (comprobante is null) return NotFound();

        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, id);
        if (comprobanteBase is null) return NotFound();

        var codigoDocumento = (comprobanteBase.CodigoDocumentoComprobante ?? string.Empty).Trim();
        var esBoleta = string.Equals(codigoDocumento, "03", StringComparison.OrdinalIgnoreCase);
        var estadoPermitido = comprobanteBase.Estado == SistemaControlEspaciosDeportivosWeb.Models.EstadoComprobanteElectronico.AceptadoSunat
                              || (esBoleta && comprobanteBase.Estado == SistemaControlEspaciosDeportivosWeb.Models.EstadoComprobanteElectronico.EnviadoSunat);
        if (!estadoPermitido)
        {
            TempData["ComprobanteError"] = "Solo se puede enviar por correo un comprobante aceptado por SUNAT o una boleta enviada a SUNAT.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        var correoDestino = (comprobante.ClienteCorreo ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(correoDestino))
        {
            TempData["ComprobanteError"] = "El cliente no tiene correo registrado para enviar el comprobante.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        var urls = await ObtenerUrlsDocumentosSunatAsync(negocioId, id);
        var adjuntos = ConstruirAdjuntosCorreo(comprobante, urls);

        if (adjuntos.Count == 0)
        {
            TempData["ComprobanteError"] = "No hay URLs PDF/XML/CDR disponibles para adjuntar en el correo.";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        var asunto = $"Comprobante electronico {comprobante.Serie}-{comprobante.Numero:D8}";
        var html = BuildComprobanteEmailHtml(comprobante, urls);
        try
        {
            await emailService.SendEmailAsync(
                correoDestino,
                string.IsNullOrWhiteSpace(comprobante.ClienteNombre) ? correoDestino : comprobante.ClienteNombre,
                asunto,
                html,
                new EmailSendOptions
                {
                    SenderEmail = "info@lazonadeportiva.com",
                    SenderName = "La Zona Deportiva",
                    AttachmentUrls = adjuntos
                });
        }
        catch (EmailDeliveryException ex)
        {
            TempData["ComprobanteError"] = $"No se pudo enviar el correo: {ex.Message}";
            return RedirectToAction(nameof(Preview), new { negocioId, id });
        }

        TempData["ComprobanteOk"] = $"Comprobante enviado por correo a {correoDestino}.";
        return RedirectToAction(nameof(Preview), new { negocioId, id });
    }

    private async Task<(string? UrlPdfSunat, string? UrlXmlSunat, string? UrlCdrSunat)> ObtenerUrlsDocumentosSunatAsync(int negocioId, int comprobanteId)
    {
        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand(
            @"SELECT TOP (1) UrlPdfSunat, UrlXmlSunat, UrlCdrSunat
              FROM dbo.ComprobantesElectronicos
              WHERE NegocioId = @NegocioId
                AND Id = @Id;", cn);
        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;
        cmd.Parameters.Add("@Id", SqlDbType.Int).Value = comprobanteId;
        await using var dr = await cmd.ExecuteReaderAsync();
        if (!await dr.ReadAsync())
            return (null, null, null);

        return (
            dr.IsDBNull(0) ? null : dr.GetString(0),
            dr.IsDBNull(1) ? null : dr.GetString(1),
            dr.IsDBNull(2) ? null : dr.GetString(2)
        );
    }

    private static List<EmailAttachmentUrlOption> ConstruirAdjuntosCorreo(
        ComprobanteVisualizacionViewModel comprobante,
        (string? UrlPdfSunat, string? UrlXmlSunat, string? UrlCdrSunat) urls)
    {
        var nombreBase = $"{comprobante.Serie}-{comprobante.Numero:D8}";
        var resultado = new List<EmailAttachmentUrlOption>();

        if (!string.IsNullOrWhiteSpace(urls.UrlPdfSunat))
        {
            resultado.Add(new EmailAttachmentUrlOption
            {
                Url = urls.UrlPdfSunat!,
                FileName = $"{nombreBase}.pdf"
            });
        }

        return resultado;
    }

    private static string BuildComprobanteEmailHtml(
        ComprobanteVisualizacionViewModel comprobante,
        (string? UrlPdfSunat, string? UrlXmlSunat, string? UrlCdrSunat) urls)
    {
        var tipoDocumento = ResolverNombreDocumento(comprobante.CodigoDocumentoComprobante, comprobante.TipoDocumentoNombre);
        var numero = $"{comprobante.Serie}-{comprobante.Numero:D8}";
        var ubigeoCliente = $"{Sanitize(comprobante.ClienteDepartamento)} / {Sanitize(comprobante.ClienteProvincia)} / {Sanitize(comprobante.ClienteDistrito)}";
        var links = new List<string>();
        if (!string.IsNullOrWhiteSpace(urls.UrlPdfSunat)) links.Add($"<li><a href=\"{EscapeUrl(urls.UrlPdfSunat!)}\">Descargar PDF</a></li>");
        if (!string.IsNullOrWhiteSpace(urls.UrlXmlSunat)) links.Add($"<li><a href=\"{EscapeUrl(urls.UrlXmlSunat!)}\">Descargar XML</a></li>");
        if (!string.IsNullOrWhiteSpace(urls.UrlCdrSunat)) links.Add($"<li><a href=\"{EscapeUrl(urls.UrlCdrSunat!)}\">Descargar CDR</a></li>");

        var linksHtml = links.Count > 0 ? string.Join(Environment.NewLine, links) : "<li>No se registraron URLs de descarga.</li>";

        return $"""
<!doctype html>
<html lang="es">
  <body style="margin:0;padding:0;background-color:#f4f7fb;font-family:Manrope,'Segoe UI',Arial,sans-serif;color:#1f2937;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color:#f4f7fb;padding:24px 12px;">
      <tr>
        <td align="center">
          <table role="presentation" width="680" cellspacing="0" cellpadding="0" style="max-width:680px;background:#ffffff;border:1px solid #dbe6f4;border-radius:14px;overflow:hidden;">
            <tr>
              <td style="background:linear-gradient(135deg,#0d3b66 0%,#164f86 60%,#17a2b8 100%);padding:22px 24px;">
                <div style="font-size:12px;letter-spacing:.08em;text-transform:uppercase;color:#dbeafe;font-weight:800;">La Zona Deportiva</div>
                <h1 style="margin:8px 0 0;font-size:26px;line-height:1.15;color:#ffffff;">Comprobante electronico</h1>
              </td>
            </tr>
            <tr>
              <td style="padding:24px;">
                <p style="margin:0 0 8px;font-size:16px;line-height:1.6;color:#334155;">Hola {Escape(Sanitize(comprobante.ClienteNombre))},</p>
                <p style="margin:0 0 12px;font-size:14px;line-height:1.6;color:#475569;">Te enviamos tu comprobante tributario aceptado por SUNAT.</p>
                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-top:8px;border:1px solid #dbe6f4;border-radius:10px;overflow:hidden;">
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">Complejo deportivo</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(Sanitize(comprobante.NegocioNombre))}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;width:220px;">Tipo de documento</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(tipoDocumento)}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;">Numero</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(numero)}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;">Fecha de emision</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{comprobante.FechaEmision:dd/MM/yyyy}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;">Cliente</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(Sanitize(comprobante.ClienteNombre))}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;">Documento cliente</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(Sanitize(comprobante.ClienteDocumento))}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;border-bottom:1px solid #e2e8f0;font-size:13px;font-weight:700;color:#0f172a;">Ubigeo cliente</td><td style="padding:10px 12px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#334155;">{Escape(ubigeoCliente)}</td></tr>
                  <tr><td style="padding:10px 12px;background:#f8fbff;font-size:13px;font-weight:700;color:#0f172a;">Total</td><td style="padding:10px 12px;font-size:13px;color:#334155;font-weight:700;">{Escape(comprobante.MonedaSimbolo)} {comprobante.Total:N2}</td></tr>
                </table>
                <div style="margin-top:16px;padding:12px;border:1px solid #dbe6f4;border-radius:10px;background:#f8fbff;">
                  <div style="font-size:13px;font-weight:700;color:#0f172a;margin-bottom:6px;">Documentos SUNAT</div>
                  <p style="margin:0 0 8px;color:#475569;font-size:12px;">Se adjunta el PDF. XML y CDR estan disponibles en estos enlaces:</p>
                  <ul style="margin:0;padding-left:18px;color:#334155;font-size:13px;line-height:1.6;">
                    {linksHtml}
                  </ul>
                </div>
                <p style="margin:16px 0 0;font-size:12px;line-height:1.5;color:#94a3b8;">Este correo fue enviado automaticamente por La Zona Deportiva.</p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>
""";
    }

    private static string ResolverNombreDocumento(string? codigoDoc, string? fallback)
    {
        var codigo = (codigoDoc ?? string.Empty).Trim().ToUpperInvariant();
        return codigo switch
        {
            "01" => "Factura",
            "03" => "Boleta",
            "07" => "Nota de Credito",
            "08" => "Nota de Debito",
            "RI" => "Recibo Interno",
            _ => string.IsNullOrWhiteSpace(fallback) ? "Comprobante" : fallback.Trim()
        };
    }

    private static string Escape(string value)
    {
        return WebUtility.HtmlEncode(value ?? string.Empty);
    }

    private static string EscapeUrl(string value)
    {
        return WebUtility.HtmlEncode(value ?? string.Empty);
    }

    private static string Sanitize(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();
    }

    private async Task PoblarDatosComprobanteAsync(ComprobanteFormViewModel model, ModuloBaseViewModel baseVm, int? reservaId)
    {
        var config = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        model.EmisionComprobantesElectronicos = config?.EmisionComprobantesElectronicos == true;
        model.EmisionReciboInterno = config?.EmisionReciboInterno == true;
        model.PorcentajeIgvConfigurado = config?.PorcentajeIgv ?? 18;
        if (model.Id == 0 && config is not null && config.MonedaId > 0)
        {
            model.TipoMoneda = (SistemaControlEspaciosDeportivosWeb.Models.TipoMoneda)config.MonedaId;
        }

        model.Reservas = await spService.ComprobantesBuscarReservasPagadasAsync(model.NegocioId, null, reservaId, 20);
        List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem> documentos;
        if (model.EmisionComprobantesElectronicos && model.EmisionReciboInterno)
        {
            documentos = await spService.CombosDocumentosComprobanteNegocioAsync(model.NegocioId, null);
        }
        else if (model.EmisionComprobantesElectronicos)
        {
            documentos = await spService.CombosDocumentosComprobanteNegocioAsync(model.NegocioId, true);
        }
        else if (model.EmisionReciboInterno)
        {
            documentos = await spService.CombosDocumentosComprobanteNegocioAsync(model.NegocioId, false);
        }
        else
        {
            documentos = new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        }
        if (!model.EsNota)
        {
            documentos = documentos
                .Where(d =>
                {
                    var codigo = (d.Value ?? string.Empty).Trim().ToUpperInvariant();
                    return codigo is not ("07" or "08");
                })
                .ToList();

            var restringirPorCodigo = reservaId.HasValue && reservaId.Value > 0;
            if (restringirPorCodigo)
            {
                var codigoActual = (model.CodigoDocumentoComprobante ?? string.Empty).Trim().ToUpperInvariant();
                if (codigoActual == "RI")
                {
                    documentos = documentos
                        .Where(d => string.Equals((d.Value ?? string.Empty).Trim(), "RI", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
                else if (codigoActual is "01" or "03")
                {
                    documentos = documentos
                        .Where(d =>
                        {
                            var codigo = (d.Value ?? string.Empty).Trim().ToUpperInvariant();
                            return codigo is "01" or "03";
                        })
                        .ToList();
                }
            }
        }
        model.TiposDocumentoComprobante = documentos;
        if (string.IsNullOrWhiteSpace(model.CodigoDocumentoComprobante))
        {
            model.CodigoDocumentoComprobante = documentos.FirstOrDefault()?.Value ?? "03";
        }

        model.DocumentoTributario = model.CodigoDocumentoComprobante != "RI";
        model.TiposDocumentoIdentidad = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        model.SeriesDocumento = string.IsNullOrWhiteSpace(model.CodigoDocumentoComprobante)
            ? new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>()
            : await spService.CombosSeriesDocumentoComprobanteAsync(model.NegocioId, model.CodigoDocumentoComprobante);

        if (model.NegocioSerieId.HasValue && !model.SeriesDocumento.Any(x => x.Value == model.NegocioSerieId.Value.ToString()))
        {
            model.NegocioSerieId = null;
        }

        if (model.NegocioSerieId.HasValue)
        {
            var serieElegida = model.SeriesDocumento.FirstOrDefault(x => x.Value == model.NegocioSerieId.Value.ToString());
            if (serieElegida is not null) model.Serie = serieElegida.Text;
        }

        if (model.ReservaId > 0)
        {
            var ctx = await spService.ComprobantesObtenerContextoReservaAsync(model.NegocioId, model.ReservaId, model.CodigoDocumentoComprobante);
            if (ctx is not null)
            {
                model.ReservaCodigo = ctx.ReservaCodigo;
                model.Sede = ctx.Sede;
                model.Espacio = ctx.Espacio;
                model.Cliente = ctx.Cliente;
                model.ClienteId = ctx.ClienteId;
                if (string.IsNullOrWhiteSpace(model.ClienteCorreo)) model.ClienteCorreo = ctx.ClienteCorreo;
                if (string.IsNullOrWhiteSpace(model.ClienteTipoDocumento)) model.ClienteTipoDocumento = ctx.ClienteTipoDocumento;
                if (string.IsNullOrWhiteSpace(model.ClienteNumeroDocumento)) model.ClienteNumeroDocumento = ctx.ClienteNumeroDocumento;
                if (string.IsNullOrWhiteSpace(model.ClienteDireccionFiscal)) model.ClienteDireccionFiscal = ctx.ClienteDireccionFiscal;
                if (string.IsNullOrWhiteSpace(model.ClienteCodigoUbigeo)) model.ClienteCodigoUbigeo = ctx.ClienteCodigoUbigeo;
                if (string.IsNullOrWhiteSpace(model.ClienteCodigoDepartamento)) model.ClienteCodigoDepartamento = ctx.ClienteCodigoDepartamento;
                if (string.IsNullOrWhiteSpace(model.ClienteCodigoProvincia)) model.ClienteCodigoProvincia = ctx.ClienteCodigoProvincia;
                model.FechaReserva = ctx.FechaReserva;
                model.HoraInicioReserva = ctx.HoraInicioReserva;
                model.HoraFinReserva = ctx.HoraFinReserva;
                model.TotalReserva = ctx.TotalReserva;
                model.PagadoReserva = ctx.TotalPagado;
                model.SaldoReserva = ctx.SaldoPendiente;
                model.MonedaSimbolo = ctx.MonedaSimbolo;
                model.PagosReserva = ctx.PagosReserva;
                model.SeriesDocumento = ctx.SeriesDisponibles;
                if (model.NegocioSerieId.HasValue
                    && !model.SeriesDocumento.Any(s => s.Value == model.NegocioSerieId.Value.ToString(CultureInfo.InvariantCulture)))
                {
                    model.NegocioSerieId = null;
                    model.Serie = string.Empty;
                }
                if (model.Total <= 0) model.Total = ctx.TotalReserva;
                if (!model.NegocioSerieId.HasValue)
                {
                    var primeraSerie = ctx.SeriesDisponibles.FirstOrDefault();
                    if (primeraSerie is not null)
                    {
                        model.NegocioSerieId = int.TryParse(primeraSerie.Value, out var idSerie) ? idSerie : null;
                        model.Serie = primeraSerie.Text;
                    }
                }
            }
            else
            {
                model.ReservaId = 0;
                model.ReservaCodigo = null;
                model.Sede = null;
                model.Espacio = null;
                model.Cliente = null;
                model.ClienteId = null;
                model.ClienteCorreo = null;
                model.ClienteTipoDocumento = null;
                model.ClienteNumeroDocumento = null;
                model.ClienteDireccionFiscal = null;
                model.ClienteCodigoDepartamento = null;
                model.ClienteCodigoProvincia = null;
                model.ClienteCodigoUbigeo = null;
                model.FechaReserva = null;
                model.HoraInicioReserva = null;
                model.HoraFinReserva = null;
                model.TotalReserva = null;
                model.PagadoReserva = null;
                model.SaldoReserva = null;
                model.PagosReserva = new List<PagoPrevioItemViewModel>();
                model.NegocioSerieId = null;
                model.Serie = string.Empty;
                if (model.Id == 0)
                {
                    model.Total = 0m;
                    model.SubTotal = 0m;
                    model.Igv = 0m;
                }
            }
        }
        else
        {
            model.ReservaCodigo = null;
            model.Sede = null;
            model.Espacio = null;
            model.Cliente = null;
            model.ClienteId = null;
            model.ClienteCorreo = null;
            model.ClienteTipoDocumento = null;
            model.ClienteNumeroDocumento = null;
            model.ClienteDireccionFiscal = null;
            model.ClienteCodigoDepartamento = null;
            model.ClienteCodigoProvincia = null;
            model.ClienteCodigoUbigeo = null;
            model.FechaReserva = null;
            model.HoraInicioReserva = null;
            model.HoraFinReserva = null;
            model.TotalReserva = null;
            model.PagadoReserva = null;
            model.SaldoReserva = null;
            model.PagosReserva = new List<PagoPrevioItemViewModel>();
        }

        if (!string.IsNullOrWhiteSpace(model.ClienteCodigoUbigeo) && Regex.IsMatch(model.ClienteCodigoUbigeo, @"^\d{6}$"))
        {
            model.ClienteCodigoDepartamento = model.ClienteCodigoUbigeo[..2];
            model.ClienteCodigoProvincia = model.ClienteCodigoUbigeo[..4];
        }

        model.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(model.ClienteCodigoDepartamento) && model.ClienteCodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(model.ClienteCodigoDepartamento)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        model.DistritosUbigeo = !string.IsNullOrWhiteSpace(model.ClienteCodigoProvincia) && model.ClienteCodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(model.ClienteCodigoProvincia)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
    }

    private void AplicarCalculoComprobante(ComprobanteFormViewModel model)
    {
        if (model.Total < 0) model.Total = 0;
        var porcentaje = model.PorcentajeIgvConfigurado < 0 ? 0 : model.PorcentajeIgvConfigurado;
        model.DocumentoTributario = model.CodigoDocumentoComprobante != "RI";
        if (!model.DocumentoTributario)
        {
            model.SubTotal = model.Total;
            model.Igv = 0m;
            return;
        }

        var divisor = 1m + (porcentaje / 100m);
        if (divisor <= 0) divisor = 1m;
        model.SubTotal = decimal.Round(model.Total / divisor, 2, MidpointRounding.AwayFromZero);
        model.Igv = decimal.Round(model.Total - model.SubTotal, 2, MidpointRounding.AwayFromZero);
    }

    private static SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante MapearTipoComprobante(string? codigoSunat)
    {
        var codigo = (codigoSunat ?? string.Empty).Trim().ToUpperInvariant();
        return codigo switch
        {
            "01" => SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante.Factura,
            "07" => SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante.NotaCredito,
            "08" => SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante.NotaDebito,
            "RI" => (SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante)3,
            _ => SistemaControlEspaciosDeportivosWeb.Models.TipoComprobante.Boleta
        };
    }

    private static string? NormalizarTipoNota(string? tipoNota)
    {
        var tipo = (tipoNota ?? string.Empty).Trim().ToUpperInvariant();
        if (tipo is "NC" or "ND") return tipo;
        return null;
    }

    private static string CodigoDocumentoPorTipoNota(string tipoNota)
    {
        return string.Equals(tipoNota, "ND", StringComparison.OrdinalIgnoreCase) ? "08" : "07";
    }

    private static string NombreDocumentoPorCodigo(string? codigoDocumento)
    {
        var codigo = (codigoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        return codigo switch
        {
            "01" => "Factura",
            "03" => "Boleta",
            "07" => "Nota de credito",
            "08" => "Nota de debito",
            "RI" => "Recibo Interno",
            _ => "Comprobante"
        };
    }

    private static bool EsComprobanteAptoParaNota(ComprobanteFormViewModel comprobante)
    {
        if (comprobante.Estado != SistemaControlEspaciosDeportivosWeb.Models.EstadoComprobanteElectronico.AceptadoSunat)
            return false;

        var codigo = (comprobante.CodigoDocumentoComprobante ?? string.Empty).Trim().ToUpperInvariant();
        return codigo is "01" or "03";
    }

    private void NormalizarDatosCliente(ComprobanteFormViewModel model)
    {
        model.ClienteCorreo = string.IsNullOrWhiteSpace(model.ClienteCorreo) ? null : model.ClienteCorreo.Trim();
        model.ClienteTipoDocumento = string.IsNullOrWhiteSpace(model.ClienteTipoDocumento) ? null : model.ClienteTipoDocumento.Trim();
        model.ClienteNumeroDocumento = string.IsNullOrWhiteSpace(model.ClienteNumeroDocumento) ? null : model.ClienteNumeroDocumento.Trim();
        model.ClienteDireccionFiscal = string.IsNullOrWhiteSpace(model.ClienteDireccionFiscal) ? null : model.ClienteDireccionFiscal.Trim();
        model.ClienteCodigoDepartamento = string.IsNullOrWhiteSpace(model.ClienteCodigoDepartamento) ? null : model.ClienteCodigoDepartamento.Trim();
        model.ClienteCodigoProvincia = string.IsNullOrWhiteSpace(model.ClienteCodigoProvincia) ? null : model.ClienteCodigoProvincia.Trim();
        model.ClienteCodigoUbigeo = string.IsNullOrWhiteSpace(model.ClienteCodigoUbigeo) ? null : model.ClienteCodigoUbigeo.Trim();
    }

    private async Task<decimal> ObtenerMontoMaximoBoletaSinDocAsync()
    {
        const string nombreParametro = "VALIDA_MONTO_BSINDOC";
        var valorParametro = await spService.ParametrosGlobalesObtenerValorAsync(nombreParametro);
        if (decimal.TryParse(valorParametro, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) && valor >= 0m)
            return valor;
        if (decimal.TryParse(valorParametro, NumberStyles.Any, CultureInfo.CurrentCulture, out valor) && valor >= 0m)
            return valor;
        return 700m;
    }

    private void ValidarReglasDocumento(ComprobanteFormViewModel model, decimal montoMaximoBoletaSinDoc)
    {
        var codigoDocComprobante = (model.CodigoDocumentoComprobante ?? string.Empty).Trim().ToUpperInvariant();
        var tipoDocCliente = (model.ClienteTipoDocumento ?? string.Empty).Trim().ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(tipoDocCliente))
            ModelState.AddModelError(nameof(model.ClienteTipoDocumento), "Selecciona el tipo de documento del cliente.");

        if (codigoDocComprobante == "03")
        {
            if (model.Total > montoMaximoBoletaSinDoc)
            {
                if (tipoDocCliente is not ("1" or "6"))
                    ModelState.AddModelError(nameof(model.ClienteTipoDocumento), $"Si la boleta supera S/ {montoMaximoBoletaSinDoc:0.00}, el tipo de documento debe ser DNI (1) o RUC (6).");

                if (string.IsNullOrWhiteSpace(model.ClienteNumeroDocumento))
                    ModelState.AddModelError(nameof(model.ClienteNumeroDocumento), $"Si la boleta supera S/ {montoMaximoBoletaSinDoc:0.00}, el numero de documento es obligatorio.");
            }
        }
        else if (codigoDocComprobante == "01")
        {
            if (tipoDocCliente != "6")
                ModelState.AddModelError(nameof(model.ClienteTipoDocumento), "Para factura el cliente debe tener RUC (6).");

            if (string.IsNullOrWhiteSpace(model.ClienteNumeroDocumento))
            {
                ModelState.AddModelError(nameof(model.ClienteNumeroDocumento), "Para factura el numero de documento (RUC) es obligatorio.");
            }
            else
            {
                var numeroDocumento = model.ClienteNumeroDocumento.Trim();
                if (!Regex.IsMatch(numeroDocumento, "^\\d{11}$"))
                    ModelState.AddModelError(nameof(model.ClienteNumeroDocumento), "Para factura, el RUC del cliente debe tener exactamente 11 digitos numericos.");
            }
        }
    }

    private void ValidarSerieNotaSegunReferencia(ComprobanteFormViewModel model, ComprobanteFormViewModel? referencia)
    {
        if (!model.EsNota || referencia is null)
            return;

        var codigoReferencia = (referencia.CodigoDocumentoComprobante ?? string.Empty).Trim().ToUpperInvariant();
        if (codigoReferencia is not ("01" or "03"))
            return;

        var serieNota = (model.Serie ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(serieNota))
            return;

        var prefijoEsperado = codigoReferencia == "01" ? 'F' : 'B';
        if (serieNota[0] != prefijoEsperado)
        {
            var tipoRef = codigoReferencia == "01" ? "Factura" : "Boleta";
            ModelState.AddModelError(
                nameof(model.NegocioSerieId),
                $"La serie de la nota debe iniciar con '{prefijoEsperado}' porque el comprobante de referencia es {tipoRef}.");
        }
    }
}
