using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Globalization;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ComprobantesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
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
            TiposDocumentoFiltro = tiposDocumentoFiltro,
            Comprobantes = comprobantes
        };
        return View(vm);
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

        int id;
        try
        {
            id = await spService.ComprobantesCrearAsync(model, User.Identity?.Name ?? "sistema");
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
        var comprobanteCreado = await spService.ComprobantesObtenerAsync(model.NegocioId, id);
        if (comprobanteCreado is not null)
        {
            TempData["ComprobanteOk"] = $"Comprobante emitido: {comprobanteCreado.Serie}-{comprobanteCreado.Numero:D8}.";
        }
        return RedirectToAction(nameof(Edit), new { negocioId = model.NegocioId, id });
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
            EsNota = true,
            TipoNota = tipoNotaNormalizado,
            ComprobanteReferenciaId = referencia.Id,
            ComprobanteReferenciaTipo = NombreDocumentoPorCodigo(referencia.CodigoDocumentoComprobante),
            ComprobanteReferenciaSerie = referencia.Serie,
            ComprobanteReferenciaNumero = referencia.Numero
        };

        await PoblarDatosComprobanteAsync(vm, baseVm, vm.ReservaId);
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

        AplicarCalculoComprobante(model);
        NormalizarDatosCliente(model);
        var montoMaximoBoletaSinDoc = await ObtenerMontoMaximoBoletaSinDocAsync();
        ValidarReglasDocumento(model, montoMaximoBoletaSinDoc);
        if (!ModelState.IsValid) return View(model);

        int idGenerado;
        try
        {
            idGenerado = await spService.ComprobantesCrearAsync(model, User.Identity?.Name ?? "sistema");
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
        var comprobanteCreado = await spService.ComprobantesObtenerAsync(model.NegocioId, idGenerado);
        if (comprobanteCreado is not null)
        {
            TempData["ComprobanteOk"] = $"{(tipoNotaNormalizado == "NC" ? "Nota de credito" : "Nota de debito")} generada: {comprobanteCreado.Serie}-{comprobanteCreado.Numero:D8}.";
        }

        return RedirectToAction(nameof(Edit), new { negocioId = model.NegocioId, id = idGenerado });
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

        var ok = await spService.ComprobantesActualizarAsync(model, User.Identity?.Name ?? "sistema");
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
    public async Task<IActionResult> ObtenerContextoReserva(int negocioId, int reservaId, string? codigoSunat = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var config = await spService.ConfiguracionClubObtenerAsync(negocioId);
        var data = await spService.ComprobantesObtenerContextoReservaAsync(negocioId, reservaId, codigoSunat);
        if (data is null) return Json(new { ok = false, mensaje = "No se encontro la reserva pagada seleccionada." });

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
                documentos = data.DocumentosDisponibles.Select(x => new { value = x.Value, text = x.Text }).ToList(),
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
    public async Task<IActionResult> Visualizar(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var data = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (data is null) return NotFound();

        if (data.EsTributario)
        {
            var url = (data.UrlDescargaProveedor ?? string.Empty).Trim();
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
                (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp))
            {
                return Redirect(url);
            }

            TempData["ComprobanteError"] = "El comprobante tributario no tiene URL de descarga del proveedor.";
            return RedirectToAction(nameof(Index), new { negocioId });
        }

        var pdf = await InternalReceiptHtmlPdfBuilder.BuildAsync(data);
        var fileName = $"ReciboInterno_{data.Serie}-{data.Numero.ToString("D8", CultureInfo.InvariantCulture)}.pdf";
        return File(pdf, "application/pdf", fileName);
    }

    [HttpGet]
    public async Task<IActionResult> Preview(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "COMPROBANTES");
        if (baseVm is null) return Forbid();

        var data = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, id);
        if (data is null) return NotFound();

        return View(data);
    }

    private async Task PoblarDatosComprobanteAsync(ComprobanteFormViewModel model, ModuloBaseViewModel baseVm, int? reservaId)
    {
        var config = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        model.EmisionComprobantesElectronicos = config?.EmisionComprobantesElectronicos == true;
        model.EmisionReciboInterno = config?.EmisionReciboInterno == true;
        model.PorcentajeIgvConfigurado = config?.PorcentajeIgv ?? 18;

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
                ModelState.AddModelError(nameof(model.ClienteNumeroDocumento), "Para factura el numero de documento (RUC) es obligatorio.");
        }
    }
}
