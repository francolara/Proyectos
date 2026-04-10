using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReservasController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService,
    INotificacionEmailService notificacionEmailService)
    : ModuloControllerBase(moduloPermisoService)
{
    private const string CodigoDocumentoRucSunat = "6";
    private const string CodigoDocumentoNoDomiciliadoSinRucSunat = "0";

    public async Task<IActionResult> Index(int? negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado, DateOnly? listadoDesde, DateOnly? listadoHasta, List<int>? estadosListado, string? estadosListadoCsv = null, int paginaListado = 1)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today);
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today.AddDays(6));
        if (hasta < desde) hasta = desde;
        sedeId = AplicarSedeAsignada(baseVm, sedeId);

        var sedes = await spService.EspaciosComboSedesAsync(resolvedNegocioId.Value, baseVm.SedeIdAsignada);
        var espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value, sedeId);
        if (espacioDeportivoId.HasValue && !espacios.Any(x => x.Value == espacioDeportivoId.Value.ToString()))
            espacioDeportivoId = null;
        var tiposDocumentoClientes = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        var formasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        var configClub = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var filtroListadoDesde = listadoDesde ?? desde;
        var filtroListadoHasta = listadoHasta ?? hasta;
        if (filtroListadoHasta < filtroListadoDesde) filtroListadoHasta = filtroListadoDesde;
        var estadosSeleccionados = estadosListado ?? new List<int>();
        if (estadosSeleccionados.Count == 0 && !string.IsNullOrWhiteSpace(estadosListadoCsv))
        {
            estadosSeleccionados = estadosListadoCsv
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(x => int.TryParse(x, out var v) ? v : 0)
                .Where(x => x > 0)
                .ToList();
        }

        var estadosListadoLimpios = estadosSeleccionados
            .Where(x => x is 1 or 2 or 4 or 5 or 6)
            .Distinct()
            .OrderBy(x => x)
            .ToList();
        var estadosListadoCsvNormalizado = estadosListadoLimpios.Count == 0 ? null : string.Join(",", estadosListadoLimpios);
        const int tamanoPaginaListado = 20;
        var paginaActualListado = paginaListado < 1 ? 1 : paginaListado;
        var (reservas, totalReservasListado) = await spService.ReservasListarAsync(
            resolvedNegocioId.Value,
            filtroListadoDesde,
            filtroListadoHasta,
            sedeId,
            espacioDeportivoId,
            null,
            estadosListadoCsvNormalizado,
            paginaActualListado,
            tamanoPaginaListado);
        var resumenListadoGlobal = await spService.ReservasListadoResumenAsync(
            resolvedNegocioId.Value,
            filtroListadoDesde,
            filtroListadoHasta,
            sedeId,
            espacioDeportivoId,
            null,
            estadosListadoCsvNormalizado);
        var totalPaginasListado = Math.Max(1, (int)Math.Ceiling(totalReservasListado / (double)tamanoPaginaListado));
        if (paginaActualListado > totalPaginasListado)
        {
            paginaActualListado = totalPaginasListado;
            (reservas, totalReservasListado) = await spService.ReservasListarAsync(
                resolvedNegocioId.Value,
                filtroListadoDesde,
                filtroListadoHasta,
                sedeId,
                espacioDeportivoId,
                null,
                estadosListadoCsvNormalizado,
                paginaActualListado,
                tamanoPaginaListado);
        }

        var vm = new ReservasIndexViewModel
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
            FechaDesde = desde,
            FechaHasta = hasta,
            ListadoFechaDesde = filtroListadoDesde,
            ListadoFechaHasta = filtroListadoHasta,
            SedeId = sedeId,
            EspacioDeportivoId = espacioDeportivoId,
            Estado = estado,
            PaginaListado = paginaActualListado,
            TamanoPaginaListado = tamanoPaginaListado,
            TotalReservasListado = totalReservasListado,
            TotalPendientesListadoGlobal = resumenListadoGlobal.TotalPendientes,
            TotalPagadasListadoGlobal = resumenListadoGlobal.TotalPagadas,
            SaldoTotalListadoGlobal = resumenListadoGlobal.SaldoTotal,
            TotalPaginasListado = totalPaginasListado,
            EstadosListadoSeleccionados = estadosListadoLimpios,
            SedesFiltro = sedes,
            EspaciosFiltro = espacios,
            ClientesFiltro = new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>(),
            TiposDocumentoClientesFiltro = tiposDocumentoClientes,
            FormasPagoFiltro = formasPago,
            EstadosFiltro = ObtenerEstadosFiltro(),
            Reservas = reservas,
            Bloqueos = await spService.BloqueosListarAsync(resolvedNegocioId.Value, desde, hasta, sedeId, espacioDeportivoId),
            BloqueoForm = new BloqueoHorarioFormViewModel
            {
                NegocioId = resolvedNegocioId.Value,
                Fecha = desde,
                Espacios = espacios
            }
        };
        vm.PoliticaConfirmacionPago = configClub?.PoliticaConfirmacionPago ?? 0;
        vm.PorcentajeAdelantoMinimo = configClub?.PorcentajeAdelantoMinimo;
        vm.MonedaNombre = "PEN";
        vm.MonedaSimbolo = "S/";

        if (sedeId.HasValue)
        {
            var sedeConfig = await spService.SedesObtenerAsync(resolvedNegocioId.Value, sedeId.Value);
            if (sedeConfig is not null)
            {
                vm.CalendarioUsaHorarioSede = true;
                vm.AtiendeLunes = sedeConfig.AtiendeLunes;
                vm.AtiendeMartes = sedeConfig.AtiendeMartes;
                vm.AtiendeMiercoles = sedeConfig.AtiendeMiercoles;
                vm.AtiendeJueves = sedeConfig.AtiendeJueves;
                vm.AtiendeViernes = sedeConfig.AtiendeViernes;
                vm.AtiendeSabado = sedeConfig.AtiendeSabado;
                vm.AtiendeDomingo = sedeConfig.AtiendeDomingo;
                vm.HoraApertura = sedeConfig.HoraApertura;
                vm.HoraCierre = sedeConfig.HoraCierre;
                vm.FechasNoLaborables = sedeConfig.FechasInhabilitadas
                    .Select(x => x.ToString("yyyy-MM-dd"))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();
            }
        }
        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerEspaciosFiltro(int negocioId, int? sedeId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return Json(new { ok = false, mensaje = "No autorizado." });

        sedeId = AplicarSedeAsignada(baseVm, sedeId);
        if (!sedeId.HasValue)
            return Json(new { ok = true, items = Array.Empty<object>() });

        var espacios = await spService.ReservasComboEspaciosAsync(negocioId, sedeId);
        return Json(new
        {
            ok = true,
            items = espacios.Select(x => new { value = x.Value, text = x.Text })
        });
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerClientesFiltro(int negocioId, string? buscar = null, int? clienteId = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return Json(new { ok = false, mensaje = "No autorizado." });

        var criterio = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var clientes = await spService.ReservasBuscarClientesAsync(negocioId, criterio, clienteId, 30);
        var idsClientes = clientes
            .Select(x => int.TryParse(x.Value, out var id) ? id : 0)
            .Where(x => x > 0)
            .Distinct()
            .ToList();

        var detalleClientes = new Dictionary<int, ClienteFormViewModel>();
        foreach (var id in idsClientes)
        {
            var detalle = await spService.ClientesObtenerAsync(negocioId, id);
            if (detalle is not null)
                detalleClientes[id] = detalle;
        }

        return Json(new
        {
            ok = true,
            items = clientes.Select(x =>
            {
                var id = int.TryParse(x.Value, out var parsed) ? parsed : 0;
                detalleClientes.TryGetValue(id, out var detalle);
                return new
                {
                    value = x.Value,
                    text = x.Text,
                    numero = string.IsNullOrWhiteSpace(detalle?.Telefono) ? string.Empty : detalle!.Telefono,
                    correo = string.IsNullOrWhiteSpace(detalle?.Correo) ? string.Empty : detalle!.Correo
                };
            })
        });
    }

    [HttpPost]
    public async Task<IActionResult> CrearClienteRapido([FromBody] ReservaClienteRapidoRequestViewModel request)
    {
        var baseVm = await ObtenerBaseAsync(request.NegocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje) || !baseVm.PuedeCrear)
            return Forbid();

        var tipoDocumento = (request.TipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(tipoDocumento))
            return BadRequest(new { ok = false, mensaje = "Selecciona el tipo de documento." });

        var numeroDocumento = string.IsNullOrWhiteSpace(request.NumeroDocumento) ? string.Empty : request.NumeroDocumento.Trim();
        var nombres = string.IsNullOrWhiteSpace(request.Nombres) ? null : request.Nombres.Trim();
        var apellidos = string.IsNullOrWhiteSpace(request.Apellidos) ? null : request.Apellidos.Trim();
        var razonSocial = string.IsNullOrWhiteSpace(request.RazonSocial) ? null : request.RazonSocial.Trim();
        var nombreEquipo = string.IsNullOrWhiteSpace(request.NombreEquipo) ? null : request.NombreEquipo.Trim();
        var telefono = string.IsNullOrWhiteSpace(request.Telefono) ? null : request.Telefono.Trim();
        var correo = string.IsNullOrWhiteSpace(request.Correo) ? null : request.Correo.Trim();

        var esRuc = tipoDocumento == CodigoDocumentoRucSunat || tipoDocumento == "RUC";
        var esNoDomiciliadoSinRuc = tipoDocumento == CodigoDocumentoNoDomiciliadoSinRucSunat;

        if (esNoDomiciliadoSinRuc)
        {
            numeroDocumento = string.Empty;
        }
        else
        {
            if (string.IsNullOrWhiteSpace(numeroDocumento))
                return BadRequest(new { ok = false, mensaje = "Ingresa el numero de documento." });
            if (numeroDocumento.Length > 11 || !numeroDocumento.All(char.IsDigit))
                return BadRequest(new { ok = false, mensaje = "El numero de documento debe tener maximo 11 digitos numericos." });
        }

        string nombresORazonSocial;
        if (esRuc)
        {
            if (string.IsNullOrWhiteSpace(razonSocial))
                return BadRequest(new { ok = false, mensaje = "Ingresa la razon social." });
            nombresORazonSocial = razonSocial;
            nombres = null;
            apellidos = null;
        }
        else
        {
            if (string.IsNullOrWhiteSpace(nombres) || string.IsNullOrWhiteSpace(apellidos))
                return BadRequest(new { ok = false, mensaje = "Ingresa nombres y apellidos." });
            nombresORazonSocial = $"{nombres} {apellidos}".Trim();
        }

        var cliente = new ClienteFormViewModel
        {
            NegocioId = request.NegocioId,
            TipoDocumento = tipoDocumento,
            NumeroDocumento = numeroDocumento,
            Nombres = nombres,
            Apellidos = apellidos,
            NombresORazonSocial = nombresORazonSocial,
            NombreEquipo = nombreEquipo,
            Telefono = telefono,
            Correo = correo,
            DireccionFiscal = null,
            CodigoUbigeo = null,
            Activo = true
        };

        try
        {
            var id = await spService.ClientesCrearAsync(cliente, User.Identity?.Name ?? "sistema");
            var etiqueta = $"{nombresORazonSocial} ({numeroDocumento})";
            if (!string.IsNullOrWhiteSpace(nombreEquipo))
                etiqueta += $" - Equipo: {nombreEquipo}";

            return Json(new
            {
                ok = true,
                clienteId = id,
                clienteTexto = etiqueta,
                numero = string.IsNullOrWhiteSpace(telefono) ? string.Empty : telefono,
                correo = string.IsNullOrWhiteSpace(correo) ? string.Empty : correo
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerReservaModal(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        var vm = await spService.ReservasObtenerAsync(negocioId, id);
        if (vm is null) return NotFound(new { ok = false, mensaje = "No se encontro la reserva." });
        if (!await EspacioPermitidoAsync(baseVm, negocioId, vm.EspacioDeportivoId))
            return Forbid();

        return Json(new
        {
            ok = true,
            id = vm.Id,
            espacioDeportivoId = vm.EspacioDeportivoId,
            clienteId = vm.ClienteId,
            fecha = vm.Fecha.ToString("yyyy-MM-dd"),
            horaInicio = vm.HoraInicio.ToString("HH\\:mm"),
            horaFin = vm.HoraFin.ToString("HH\\:mm"),
            estado = (int)vm.Estado,
            total = vm.Total,
            adelanto = vm.Adelanto,
            comentario = vm.Comentario
        });
    }

    [HttpGet]
    public async Task<IActionResult> CotizarReserva(int negocioId, int espacioDeportivoId, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();
        if (!await EspacioPermitidoAsync(baseVm, negocioId, espacioDeportivoId)) return Forbid();

        try
        {
            var cotizacion = await spService.ReservasCotizarAsync(negocioId, espacioDeportivoId, fecha, horaInicio, horaFin);
            return Json(new
            {
                ok = cotizacion.Ok,
                mensaje = cotizacion.Mensaje,
                precioBase = cotizacion.PrecioBase,
                descuentoPct = cotizacion.DescuentoPct,
                precioFinal = cotizacion.PrecioFinal,
                monedaSimbolo = cotizacion.MonedaSimbolo,
                monedaNombre = cotizacion.MonedaNombre,
                politicaConfirmacionPago = cotizacion.PoliticaConfirmacionPago,
                porcentajeAdelantoMinimo = cotizacion.PorcentajeAdelantoMinimo
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpGet]
    public async Task<IActionResult> ValidarDisponibilidadModal(int negocioId, int? reservaId, int espacioDeportivoId, DateOnly fecha, TimeOnly horaInicio, TimeOnly horaFin)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();
        if (EsFechaPasada(fecha))
        {
            return Json(new
            {
                ok = true,
                disponible = false,
                mensaje = "No se permite registrar reservas en fechas pasadas.",
                conflictoTipo = (string?)null,
                conflictoId = (int?)null
            });
        }
        if (!await EspacioPermitidoAsync(baseVm, negocioId, espacioDeportivoId))
        {
            return Json(new
            {
                ok = true,
                disponible = false,
                mensaje = "No tienes acceso a la sede del espacio seleccionado.",
                conflictoTipo = (string?)null,
                conflictoId = (int?)null
            });
        }

        var validacion = await spService.ReservasValidarDisponibilidadAsync(negocioId, reservaId, espacioDeportivoId, fecha, horaInicio, horaFin);
        return Json(new
        {
            ok = true,
            disponible = validacion.Disponible,
            mensaje = validacion.Mensaje,
            conflictoTipo = validacion.ConflictoTipo,
            conflictoId = validacion.ConflictoId
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarReservaModal(ReservaFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        var requiereEditar = model.Id > 0;
        var autorizado = baseVm is not null &&
                         string.IsNullOrWhiteSpace(baseVm.Mensaje) &&
                         ((requiereEditar && baseVm.PuedeEditar) || (!requiereEditar && baseVm.PuedeCrear));

        if (!autorizado) return Forbid();
        var contexto = baseVm!;
        if (!ModelState.IsValid)
        {
            var errores = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            return BadRequest(new { ok = false, mensaje = "Revisa los datos ingresados.", errores });
        }
        if (model.Total <= 0)
        {
            return BadRequest(new { ok = false, mensaje = "El precio de espacio es obligatorio y debe ser mayor que cero." });
        }
        if (model.RegistrarPago)
        {
            if (model.Adelanto <= 0)
                return BadRequest(new { ok = false, mensaje = "Debes ingresar un monto de adelanto/pago mayor que cero." });
            if (!model.FormaPagoId.HasValue || model.FormaPagoId.Value <= 0)
                return BadRequest(new { ok = false, mensaje = "Selecciona una forma de pago para registrar el adelanto/pago." });

            model.FechaPago ??= DateTime.Today;
            if (model.FechaPago.Value.Date > DateTime.Today)
                return BadRequest(new { ok = false, mensaje = "La fecha de pago no puede ser mayor al dia actual." });

            model.NumeroOperacion = string.IsNullOrWhiteSpace(model.NumeroOperacion) ? null : model.NumeroOperacion.Trim();
            if (!string.IsNullOrWhiteSpace(model.NumeroOperacion) && !model.NumeroOperacion.All(char.IsLetterOrDigit))
                return BadRequest(new { ok = false, mensaje = "El numero de operacion solo puede contener caracteres alfanumericos." });
        }
        else
        {
            model.Adelanto = 0;
            model.FormaPagoId = null;
            model.FechaPago = null;
            model.NumeroOperacion = null;
        }
        if (!requiereEditar)
        {
            model.Estado = EstadoReserva.Pendiente;
        }
        if (EsFechaPasada(model.Fecha))
        {
            return BadRequest(new { ok = false, mensaje = "No se permite registrar reservas en fechas pasadas." });
        }
        if (!await EspacioPermitidoAsync(contexto, model.NegocioId, model.EspacioDeportivoId))
        {
            return BadRequest(new { ok = false, mensaje = "No tienes acceso a la sede del espacio seleccionado." });
        }

        try
        {
            if (requiereEditar)
            {
                var ok = await spService.ReservasActualizarAsync(model, User.Identity?.Name ?? "sistema");
                if (!ok) return NotFound(new { ok = false, mensaje = "No se encontro la reserva para actualizar." });
            }
            else
            {
                model.Id = await spService.ReservasCrearAsync(model, User.Identity?.Name ?? "sistema");
            }

            return Json(new { ok = true, id = model.Id, mensaje = "Reserva guardada correctamente." });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpGet]
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public async Task<IActionResult> CalendarioEventos(int negocioId, DateTime? start, DateTime? end, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        var desde = DateOnly.FromDateTime(start?.Date ?? DateTime.Today);
        var hasta = DateOnly.FromDateTime((end?.Date ?? DateTime.Today.AddDays(7)).AddDays(-1));
        if (hasta < desde) hasta = desde;
        sedeId = AplicarSedeAsignada(baseVm, sedeId);

        var items = await spService.ReservasCalendarioEventosAsync(negocioId, desde, hasta, sedeId, espacioDeportivoId, estado);
        var eventos = items.Select(r => new
        {
            id = $"{r.TipoEvento}-{r.Id}",
            reservaId = r.TipoEvento == "RESERVA" ? r.Id : (int?)null,
            bloqueoId = r.TipoEvento == "BLOQUEO" ? r.Id : (int?)null,
            tipoEvento = r.TipoEvento,
            title = r.Titulo,
            start = new DateTime(r.Fecha.Year, r.Fecha.Month, r.Fecha.Day, r.HoraInicio.Hour, r.HoraInicio.Minute, 0),
            end = new DateTime(r.Fecha.Year, r.Fecha.Month, r.Fecha.Day, r.HoraFin.Hour, r.HoraFin.Minute, 0),
            estado = r.Estado,
            estadoCodigo = r.EstadoCodigo,
            estadoTexto = r.EstadoTexto,
            motivo = r.Motivo,
            totalReserva = r.TotalReserva,
            espacioDeportivoId = r.EspacioDeportivoId,
            backgroundColor = r.Color,
            borderColor = r.Color,
            editable = r.TipoEvento == "RESERVA" && baseVm.PuedeEditar
        });

        return Json(eventos);
    }

    [HttpGet]
    public async Task<IActionResult> HistorialReserva(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();
        if (!await ReservaPermitidaAsync(baseVm, negocioId, id)) return Forbid();

        var historial = await spService.ReservasHistorialAsync(negocioId, id);
        return Json(new
        {
            ok = true,
            items = historial.Select(x => new
            {
                fecha = x.FechaRegistro.ToLocalTime().ToString("dd/MM/yyyy"),
                accion = x.Accion,
                usuario = x.Usuario,
                detalle = x.Detalle
            })
        });
    }

    [HttpPost]
    public async Task<IActionResult> AccionMasiva([FromBody] ReservaAccionMasivaRequestViewModel request)
    {
        var baseVm = await ObtenerBaseAsync(request.NegocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje) || !baseVm.PuedeEditar) return Forbid();

        var ids = (request.ReservaIds ?? new List<int>())
            .Where(x => x > 0)
            .Distinct()
            .ToList();
        if (ids.Count == 0)
            return BadRequest(new { ok = false, mensaje = "Selecciona al menos una reserva." });

        var accion = (request.Accion ?? string.Empty).Trim().ToLowerInvariant();
        if (accion is not ("confirmar" or "noshow" or "recordatorio"))
            return BadRequest(new { ok = false, mensaje = "Accion no valida." });

        var procesadas = 0;
        var omitidas = 0;
        var errores = new List<string>();
        foreach (var reservaId in ids)
        {
            if (!await ReservaPermitidaAsync(baseVm, request.NegocioId, reservaId))
            {
                omitidas++;
                continue;
            }

            try
            {
                if (accion == "confirmar")
                {
                    var ok = await spService.ReservasCambiarEstadoRapidoAsync(request.NegocioId, reservaId, 2, User.Identity?.Name ?? "sistema");
                    if (ok) procesadas++; else omitidas++;
                    continue;
                }

                if (accion == "noshow")
                {
                    var ok = await spService.ReservasCambiarEstadoRapidoAsync(request.NegocioId, reservaId, 6, User.Identity?.Name ?? "sistema");
                    if (ok) procesadas++; else omitidas++;
                    continue;
                }

                var reserva = await spService.ReservasObtenerParaRecordatorioAsync(request.NegocioId, reservaId);
                if (reserva is null || string.IsNullOrWhiteSpace(reserva.Correo))
                {
                    omitidas++;
                    continue;
                }

                var enviado = await notificacionEmailService.EnviarRecordatorioReservaAsync(reserva);
                if (!enviado)
                {
                    omitidas++;
                    continue;
                }

                await spService.ReservasMarcarRecordatorioEnviadoAsync(request.NegocioId, reservaId, User.Identity?.Name ?? "sistema");
                procesadas++;
            }
            catch (Exception ex)
            {
                errores.Add($"Reserva #{reservaId}: {ex.Message}");
            }
        }

        return Json(new
        {
            ok = true,
            procesadas,
            omitidas,
            errores,
            mensaje = accion switch
            {
                "confirmar" => $"Reservas confirmadas: {procesadas}. Omitidas: {omitidas}.",
                "noshow" => $"Reservas marcadas como no asistio: {procesadas}. Omitidas: {omitidas}.",
                _ => $"Recordatorios enviados: {procesadas}. Omitidas: {omitidas}."
            }
        });
    }

    [HttpGet]
    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public async Task<IActionResult> ResumenDiaOperativo(int negocioId, DateOnly fecha, int? sedeId, int? espacioDeportivoId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        sedeId = AplicarSedeAsignada(baseVm, sedeId);
        if (!sedeId.HasValue || !espacioDeportivoId.HasValue)
            return BadRequest(new { ok = false, mensaje = "Debes seleccionar sede y espacio." });

        if (!await EspacioPermitidoAsync(baseVm, negocioId, espacioDeportivoId.Value))
            return Forbid();

        var eventos = await spService.ReservasCalendarioEventosAsync(negocioId, fecha, fecha, sedeId, espacioDeportivoId, null);

        static int ToMinutes(TimeOnly t) => (t.Hour * 60) + t.Minute;
        static bool Cruza((int Inicio, int Fin) a, (int Inicio, int Fin) b) => a.Inicio < b.Fin && b.Inicio < a.Fin;

        var tramosBloqueados = eventos
            .Where(e =>
            {
                var tipo = (e.TipoEvento ?? string.Empty).ToUpperInvariant();
                return tipo == "NO_ATENCION" || tipo == "BLOQUEO";
            })
            .Select(e => (Inicio: ToMinutes(e.HoraInicio), Fin: ToMinutes(e.HoraFin)))
            .Where(x => x.Fin > x.Inicio)
            .ToList();

        var ocupados = eventos
            .Where(e =>
            {
                var tipo = (e.TipoEvento ?? string.Empty).ToUpperInvariant();
                if (tipo == "NO_ATENCION" || tipo == "BLOQUEO") return true;
                if (tipo != "RESERVA") return false;
                if (!e.Estado.HasValue) return true;
                return e.Estado.Value != 5 && e.Estado.Value != 6;
            })
            .Select(e => (Inicio: ToMinutes(e.HoraInicio), Fin: ToMinutes(e.HoraFin)))
            .Where(x => x.Fin > x.Inicio)
            .ToList();

        var slotsDisponibles = new List<object>();
        const int inicioDia = 6 * 60;
        const int finDia = 24 * 60;
        for (var minuto = inicioDia; minuto < finDia; minuto += 60)
        {
            var tramo = (Inicio: minuto, Fin: minuto + 60);
            if (ocupados.Any(o => Cruza(tramo, o))) continue;
            var horaInicio = TimeOnly.FromTimeSpan(TimeSpan.FromMinutes(tramo.Inicio)).ToString("HH\\:mm");
            var horaFin = TimeOnly.FromTimeSpan(TimeSpan.FromMinutes(tramo.Fin)).ToString("HH\\:mm");
            slotsDisponibles.Add(new
            {
                horaInicio,
                horaFin
            });
        }

        var pendientes = eventos
            .Where(e => string.Equals(e.TipoEvento, "RESERVA", StringComparison.OrdinalIgnoreCase) && e.Estado == 1)
            .OrderBy(e => e.HoraInicio)
            .Select(e => new
            {
                reservaId = e.Id,
                titulo = e.Titulo,
                horaInicio = e.HoraInicio.ToString("HH\\:mm"),
                horaFin = e.HoraFin.ToString("HH\\:mm")
            })
            .ToList();

        var reservasActivas = eventos.Count(e =>
            string.Equals(e.TipoEvento, "RESERVA", StringComparison.OrdinalIgnoreCase)
            && e.Estado is 1 or 2 or 3 or 4);
        var bloqueosActivos = eventos.Count(e =>
            string.Equals(e.TipoEvento, "BLOQUEO", StringComparison.OrdinalIgnoreCase)
            || string.Equals(e.TipoEvento, "NO_ATENCION", StringComparison.OrdinalIgnoreCase));
        var totalSlots = 0;
        for (var minuto = inicioDia; minuto < finDia; minuto += 60)
        {
            var tramo = (Inicio: minuto, Fin: minuto + 60);
            if (tramosBloqueados.Any(b => Cruza(tramo, b))) continue;
            totalSlots++;
        }
        var slotsOcupados = Math.Max(0, totalSlots - slotsDisponibles.Count);
        var ocupacionPct = totalSlots <= 0 ? 0m : Math.Round((slotsOcupados * 100m) / totalSlots, 2);

        var espacios = await spService.ReservasComboEspaciosAsync(negocioId, sedeId);
        var eventosSede = await spService.ReservasCalendarioEventosAsync(negocioId, fecha, fecha, sedeId, null, null);
        var resumenPorEspacio = espacios
            .Select(espacio =>
            {
                var espacioId = int.TryParse(espacio.Value, out var idParsed) ? idParsed : 0;
                var eventosEspacio = eventosSede
                    .Where(e => e.EspacioDeportivoId == espacioId
                                && string.Equals(e.TipoEvento, "RESERVA", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var totalPendientes = eventosEspacio.Count(x => x.Estado == 1);
                var totalConfirmadas = eventosEspacio.Count(x => x.Estado == 2);
                var totalPagadas = eventosEspacio.Count(x => x.Estado is 3 or 4);
                var totalCanceladas = eventosEspacio.Count(x => x.Estado == 5);
                var totalNoShow = eventosEspacio.Count(x => x.Estado == 6);
                var total = totalPendientes + totalConfirmadas + totalPagadas + totalCanceladas + totalNoShow;
                return new
                {
                    espacioId,
                    espacio = espacio.Text,
                    total,
                    pendientes = totalPendientes,
                    confirmadas = totalConfirmadas,
                    pagadas = totalPagadas,
                    canceladas = totalCanceladas,
                    noShow = totalNoShow
                };
            })
            .OrderByDescending(x => x.total)
            .ThenBy(x => x.espacio)
            .ToList();

        return Json(new
        {
            ok = true,
            fecha = fecha.ToString("yyyy-MM-dd"),
            slotsDisponibles,
            pendientes,
            totalSlotsDisponibles = slotsDisponibles.Count,
            totalPendientes = pendientes.Count,
            kpi = new
            {
                totalSlots,
                slotsOcupados,
                slotsLibres = slotsDisponibles.Count,
                reservasActivas,
                bloqueosActivos,
                ocupacionPct
            },
            resumenEspacios = resumenPorEspacio
        });
    }

    [HttpPost]
    public async Task<IActionResult> MoverEvento([FromBody] ReservaMoverRequestViewModel request)
    {
        var baseVm = await ObtenerBaseAsync(request.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return Forbid();
        var inicioLocal = NormalizarFechaHoraLocal(request.Inicio);
        var finLocal = NormalizarFechaHoraLocal(request.Fin);

        if (inicioLocal >= finLocal)
        {
            return BadRequest(new { ok = false, mensaje = "El horario no es valido." });
        }
        if (EsFechaPasada(DateOnly.FromDateTime(inicioLocal)))
        {
            return BadRequest(new { ok = false, mensaje = "No se permite mover reservas a fechas pasadas." });
        }
        if (!await ReservaPermitidaAsync(baseVm, request.NegocioId, request.ReservaId))
            return Forbid();

        try
        {
            var ok = await spService.ReservasMoverAsync(
                request.NegocioId,
                request.ReservaId,
                DateOnly.FromDateTime(inicioLocal),
                TimeOnly.FromDateTime(inicioLocal),
                TimeOnly.FromDateTime(finLocal),
                User.Identity?.Name ?? "sistema");

            if (!ok) return NotFound(new { ok = false, mensaje = "No se encontro la reserva." });
            return Json(new { ok = true });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpPost]
    public async Task<IActionResult> CambiarEstadoRapido([FromBody] ReservaEstadoRapidoRequestViewModel request)
    {
        var baseVm = await ObtenerBaseAsync(request.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return Forbid();
        if (!await ReservaPermitidaAsync(baseVm, request.NegocioId, request.ReservaId))
            return Forbid();

        try
        {
            var ok = await spService.ReservasCambiarEstadoRapidoAsync(
                request.NegocioId,
                request.ReservaId,
                request.NuevoEstado,
                User.Identity?.Name ?? "sistema");

            if (!ok) return NotFound(new { ok = false, mensaje = "No se encontro la reserva." });
            return Json(new { ok = true });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CambiarEstadoTabla(int negocioId, int id, int nuevoEstado, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!await ReservaPermitidaAsync(baseVm, negocioId, id))
            return Forbid();

        try
        {
            await spService.ReservasCambiarEstadoRapidoAsync(negocioId, id, nuevoEstado, User.Identity?.Name ?? "sistema");
        }
        catch (Exception ex)
        {
            TempData["ReservasError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId, fechaDesde, fechaHasta, sedeId, espacioDeportivoId, estado });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CrearBloqueo(BloqueoHorarioFormViewModel model, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!await EspacioPermitidoAsync(baseVm, model.NegocioId, model.EspacioDeportivoId))
            return Forbid();
        sedeId = AplicarSedeAsignada(baseVm, sedeId);
        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today);
        var hasta = fechaHasta ?? desde.AddDays(6);
        if (hasta < desde) hasta = desde;

        if (!ModelState.IsValid)
        {
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId, fechaDesde = desde, fechaHasta = hasta, sedeId, espacioDeportivoId, estado });
        }

        try
        {
            await spService.BloqueosCrearAsync(model, User.Identity?.Name ?? "sistema");
        }
        catch (Exception ex)
        {
            TempData["ReservasError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId, fechaDesde = desde, fechaHasta = hasta, sedeId, espacioDeportivoId, estado });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarBloqueo(int negocioId, int id, DateOnly fechaDesde, DateOnly fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        await spService.BloqueosEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId, fechaDesde, fechaHasta, sedeId, espacioDeportivoId, estado });
    }

    public async Task<IActionResult> Create(int? negocioId, DateOnly? fecha, TimeOnly? horaInicio, TimeOnly? horaFin, int? espacioDeportivoId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new ReservaFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            Fecha = fecha.HasValue && !EsFechaPasada(fecha.Value) ? fecha.Value : DateOnly.FromDateTime(DateTime.Today),
            HoraInicio = horaInicio ?? new TimeOnly(18, 0),
            HoraFin = horaFin ?? (horaInicio?.AddHours(1) ?? new TimeOnly(19, 0))
        };
        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        vm.Espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value, sedeFiltro);
        vm.Clientes = await spService.ReservasComboClientesAsync(resolvedNegocioId.Value);
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        if (espacioDeportivoId.HasValue && vm.Espacios.Any(x => x.Value == espacioDeportivoId.Value.ToString()))
        {
            vm.EspacioDeportivoId = espacioDeportivoId.Value;
        }
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ReservaFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId, sedeFiltro);
        model.Clientes = await spService.ReservasComboClientesAsync(model.NegocioId);
        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        if (EsFechaPasada(model.Fecha))
        {
            ModelState.AddModelError(nameof(model.Fecha), "No se permite registrar reservas en fechas pasadas.");
        }
        if (!ModelState.IsValid) return View(model);
        if (!await EspacioPermitidoAsync(baseVm, model.NegocioId, model.EspacioDeportivoId))
        {
            ModelState.AddModelError(string.Empty, "No tienes acceso a la sede del espacio seleccionado.");
            return View(model);
        }

        try
        {
            await spService.ReservasCrearAsync(model, User.Identity?.Name ?? "sistema");
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

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ReservasObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!await EspacioPermitidoAsync(baseVm, resolvedNegocioId.Value, vm.EspacioDeportivoId))
            return Forbid();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        vm.Espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value, sedeFiltro);
        vm.Clientes = await spService.ReservasComboClientesAsync(resolvedNegocioId.Value);
        vm.FormasPago = await spService.PagosComboFormasPagoAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ReservaFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId, sedeFiltro);
        model.Clientes = await spService.ReservasComboClientesAsync(model.NegocioId);
        model.FormasPago = await spService.PagosComboFormasPagoAsync(model.NegocioId);
        if (EsFechaPasada(model.Fecha))
        {
            ModelState.AddModelError(nameof(model.Fecha), "No se permite registrar reservas en fechas pasadas.");
        }
        if (!ModelState.IsValid) return View(model);
        if (!await EspacioPermitidoAsync(baseVm, model.NegocioId, model.EspacioDeportivoId))
        {
            ModelState.AddModelError(string.Empty, "No tienes acceso a la sede del espacio seleccionado.");
            return View(model);
        }

        try
        {
            var ok = await spService.ReservasActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar la reserva. Verifica el negocio seleccionado.");
                return View(model);
            }
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!await ReservaPermitidaAsync(baseVm, negocioId, id))
            return Forbid();
        try
        {
            var ok = await spService.ReservasEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            if (!ok) return NotFound();
        }
        catch (Exception ex)
        {
            TempData["ReservasError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private static List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem> ObtenerEstadosFiltro()
    {
        return new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>
        {
            new("Todos", string.Empty),
            new("Pendiente", "1"),
            new("Confirmada", "2"),
            new("Pagada", "4"),
            new("Cancelada", "5"),
            new("No Asistio", "6")
        };
    }

    private static bool EsFechaPasada(DateOnly fecha)
    {
        return fecha < DateOnly.FromDateTime(DateTime.Today);
    }

    private static DateTime NormalizarFechaHoraLocal(DateTime value)
    {
        return value.Kind switch
        {
            DateTimeKind.Utc => value.ToLocalTime(),
            DateTimeKind.Unspecified => DateTime.SpecifyKind(value, DateTimeKind.Local),
            _ => value
        };
    }

    private async Task<bool> EspacioPermitidoAsync(ModuloBaseViewModel baseVm, int negocioId, int espacioDeportivoId)
    {
        if (baseVm.EsAdministrador || !baseVm.SedeIdAsignada.HasValue) return true;
        var espacios = await spService.ReservasComboEspaciosAsync(negocioId, baseVm.SedeIdAsignada);
        return espacios.Any(x => x.Value == espacioDeportivoId.ToString());
    }

    private async Task<bool> ReservaPermitidaAsync(ModuloBaseViewModel baseVm, int negocioId, int reservaId)
    {
        if (baseVm.EsAdministrador || !baseVm.SedeIdAsignada.HasValue) return true;
        var reserva = await spService.ReservasObtenerAsync(negocioId, reservaId);
        if (reserva is null) return false;
        return await EspacioPermitidoAsync(baseVm, negocioId, reserva.EspacioDeportivoId);
    }
}
