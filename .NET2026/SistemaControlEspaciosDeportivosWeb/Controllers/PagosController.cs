using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class PagosController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    IReservationEmailNotificationService reservationEmailNotificationService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId, string? buscar = null, DateOnly? fechaDesde = null, DateOnly? fechaHasta = null, string? preset = null, int pagina = 1)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PAGOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var (desde, hasta) = ResolverRangoFechas(fechaDesde, fechaHasta, preset);
        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var (pagos, totalRegistros) = await spService.PagosListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null), buscar, desde, hasta, paginaActual, tamanoPagina);
        var configClub = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
        {
            paginaActual = totalPaginas;
            (pagos, totalRegistros) = await spService.PagosListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null), buscar, desde, hasta, paginaActual, tamanoPagina);
        }

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
            Buscar = buscar,
            FechaDesde = desde,
            FechaHasta = hasta,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            MonedaSimbolo = pagos.FirstOrDefault()?.MonedaSimbolo ?? "S/",
            EmisionComprobantesElectronicos = configClub?.EmisionComprobantesElectronicos == true,
            EmisionReciboInterno = configClub?.EmisionReciboInterno == true,
            Pagos = pagos
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

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PAGOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new PagoFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        vm.Reservas = await spService.PagosBuscarReservasAsync(resolvedNegocioId.Value, null, null, 20);
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(PagoFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Reservas = await spService.PagosBuscarReservasAsync(model.NegocioId, null, model.ReservaId, 20);
        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        await PoblarResumenReservaPagoAsync(model);
        if (!ModelState.IsValid) return View(model);
        var estadoAnterior = await ObtenerEstadoReservaAsync(model.NegocioId, model.ReservaId);

        try
        {
            await spService.PagosCrearAsync(model, User.Identity?.Name ?? "sistema");
            await reservationEmailNotificationService.NotifyReservationConfirmedIfAppliesAsync(
                model.NegocioId,
                model.ReservaId,
                estadoAnterior);
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
        vm.AgregarNuevoPago = !vm.TieneComprobanteActivo;
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(PagoReservaEditViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var estadoActualReserva = await spService.PagosObtenerAsync(model.NegocioId, model.ReservaId);
        if (estadoActualReserva is null) return NotFound();
        if (estadoActualReserva.TieneComprobanteActivo)
        {
            estadoActualReserva.NegocioNombre = baseVm.NegocioNombre;
            estadoActualReserva.RolActual = baseVm.RolActual;
            estadoActualReserva.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
            estadoActualReserva.AgregarNuevoPago = false;
            ModelState.AddModelError(string.Empty, "La reserva ya tiene comprobante emitido. Esta pantalla queda solo en modo visualizacion.");
            return View(estadoActualReserva);
        }

        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        if (model.Pagos is null) model.Pagos = new List<PagoReservaDetalleItemViewModel>();
        AplicarSeleccionEliminacionDesdeForm(model);
        var intentoNuevoPago =
            (model.NuevoMonto.HasValue && model.NuevoMonto.Value > 0)
            || !string.IsNullOrWhiteSpace(model.NuevoNumeroOperacion)
            || !string.IsNullOrWhiteSpace(model.NuevaObservacion);

        // Solo consideramos "nuevo pago" cuando existen datos reales.
        // Esto evita bloquear eliminaciones/ediciones cuando el switch llega marcado por defecto.
        model.AgregarNuevoPago = intentoNuevoPago;
        if (!model.AgregarNuevoPago)
        {
            model.NuevaFechaPago = null;
            model.NuevoMonto = null;
            model.NuevaFormaPagoId = null;
            model.NuevoNumeroOperacion = null;
            model.NuevaObservacion = null;
        }

        if (model.AgregarNuevoPago)
        {
            if (!model.NuevoMonto.HasValue || model.NuevoMonto.Value <= 0)
                ModelState.AddModelError(string.Empty, "El monto del nuevo pago debe ser mayor que cero.");
            if (!model.NuevaFormaPagoId.HasValue || model.NuevaFormaPagoId.Value <= 0)
                ModelState.AddModelError(string.Empty, "Selecciona forma de pago para el nuevo pago.");
            if (!model.NuevaFechaPago.HasValue)
                ModelState.AddModelError(string.Empty, "Selecciona fecha de pago para el nuevo pago.");
        }
        if (!ModelState.IsValid) return View(model);
        var estadoAnterior = await ObtenerEstadoReservaAsync(model.NegocioId, model.ReservaId);

        try
        {
            var usuario = User.Identity?.Name ?? "sistema";
            foreach (var pago in model.Pagos)
            {
                if (pago.Eliminar)
                {
                    var okEliminar = await spService.PagosEliminarAsync(model.NegocioId, pago.PagoId, usuario);
                    if (!okEliminar)
                    {
                        ModelState.AddModelError(string.Empty, $"No se pudo eliminar el pago #{pago.PagoId}.");
                        return View(model);
                    }
                    continue;
                }

                var okObs = await spService.PagosActualizarAsync(model.NegocioId, pago.PagoId, pago.Observacion, usuario);
                if (!okObs)
                {
                    ModelState.AddModelError(string.Empty, $"No se pudo actualizar la observacion del pago #{pago.PagoId}.");
                    return View(model);
                }
            }

            if (model.AgregarNuevoPago && model.NuevoMonto.HasValue && model.NuevaFormaPagoId.HasValue && model.NuevaFechaPago.HasValue)
            {
                var nuevoPago = new PagoFormViewModel
                {
                    NegocioId = model.NegocioId,
                    ReservaId = model.ReservaId,
                    FechaPago = model.NuevaFechaPago.Value,
                    Monto = model.NuevoMonto.Value,
                    FormaPagoId = model.NuevaFormaPagoId.Value,
                    NumeroOperacion = string.IsNullOrWhiteSpace(model.NuevoNumeroOperacion) ? null : model.NuevoNumeroOperacion.Trim(),
                    Observacion = string.IsNullOrWhiteSpace(model.NuevaObservacion) ? null : model.NuevaObservacion.Trim()
                };
                await spService.PagosCrearAsync(nuevoPago, usuario);
                await reservationEmailNotificationService.NotifyReservationConfirmedIfAppliesAsync(
                    model.NegocioId,
                    model.ReservaId,
                    estadoAnterior);
            }

            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
    }

    private async Task<int?> ObtenerEstadoReservaAsync(int negocioId, int reservaId)
    {
        if (reservaId <= 0) return null;
        var contexto = await spService.ReservasObtenerContextoEmailAsync(negocioId, reservaId);
        return contexto?.Estado;
    }

    private void AplicarSeleccionEliminacionDesdeForm(PagoReservaEditViewModel model)
    {
        if (!Request.HasFormContentType || model.Pagos.Count == 0) return;

        for (var i = 0; i < model.Pagos.Count; i++)
        {
            var key = $"Pagos[{i}].Eliminar";
            if (!Request.Form.TryGetValue(key, out var values))
            {
                model.Pagos[i].Eliminar = false;
                continue;
            }

            model.Pagos[i].Eliminar = values.Any(v =>
                string.Equals(v, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(v, "on", StringComparison.OrdinalIgnoreCase)
                || string.Equals(v, "1", StringComparison.OrdinalIgnoreCase));
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteByReserva(int negocioId, int reservaId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PAGOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.PagosEliminarPorReservaAsync(negocioId, reservaId, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            TempData["PagosError"] = "No se pudo eliminar los pagos de la reserva seleccionada.";
        }
        else
        {
            TempData["PagosOk"] = "Se eliminaron los pagos de la reserva y la reserva fue cancelada.";
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpGet]
    public async Task<IActionResult> BuscarReservas(int negocioId, string? q = null, int? reservaId = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PAGOS");
        if (baseVm is null) return Forbid();

        var items = await spService.PagosBuscarReservasAsync(negocioId, q, reservaId, 40);
        return Json(new
        {
            ok = true,
            items = items.Select(x => new { value = x.Value, text = x.Text })
        });
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerResumenReserva(int negocioId, int reservaId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PAGOS");
        if (baseVm is null) return Forbid();

        var data = await spService.PagosObtenerAsync(negocioId, reservaId);
        if (data is null) return Json(new { ok = false, mensaje = "No se encontró la reserva seleccionada." });

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
                fechaReserva = data.FechaReserva.ToString("yyyy-MM-dd"),
                horaInicioReserva = data.HoraInicioReserva.ToString("HH\\:mm"),
                horaFinReserva = data.HoraFinReserva.ToString("HH\\:mm"),
                totalReserva = data.TotalReserva,
                totalPagado = data.TotalPagado,
                saldoPendiente = data.SaldoPendiente,
                monedaSimbolo = data.MonedaSimbolo,
                politicaConfirmacionPago = data.PoliticaConfirmacionPago,
                porcentajeAdelantoMinimo = data.PorcentajeAdelantoMinimo,
                pagos = data.Pagos.Select(p => new
                {
                    pagoId = p.PagoId,
                    fechaPago = p.FechaPago.ToString("dd/MM/yyyy"),
                    monto = p.Monto,
                    formaPago = p.FormaPagoNombre,
                    numeroOperacion = p.NumeroOperacion,
                    observacion = p.Observacion
                })
            }
        });
    }

    private async Task PoblarResumenReservaPagoAsync(PagoFormViewModel model)
    {
        if (model.ReservaId <= 0) return;
        var data = await spService.PagosObtenerAsync(model.NegocioId, model.ReservaId);
        if (data is null) return;

        model.ReservaTextoSeleccionada = data.ReservaCodigo;
        model.Sede = data.Sede;
        model.Espacio = data.Espacio;
        model.Cliente = data.Cliente;
        model.FechaReserva = data.FechaReserva;
        model.HoraInicioReserva = data.HoraInicioReserva;
        model.HoraFinReserva = data.HoraFinReserva;
        model.TotalReserva = data.TotalReserva;
        model.PagadoReserva = data.TotalPagado;
        model.SaldoReserva = data.SaldoPendiente;
        model.MonedaSimbolo = data.MonedaSimbolo;
        model.PoliticaConfirmacionPago = data.PoliticaConfirmacionPago;
        model.PorcentajeAdelantoMinimo = data.PorcentajeAdelantoMinimo;
        model.PagosPrevios = data.Pagos
            .Select(p => new PagoPrevioItemViewModel
            {
                PagoId = p.PagoId,
                FechaPago = p.FechaPago,
                Monto = p.Monto,
                FormaPago = p.FormaPagoNombre,
                NumeroOperacion = p.NumeroOperacion,
                Observacion = p.Observacion
            }).ToList();
    }
}
