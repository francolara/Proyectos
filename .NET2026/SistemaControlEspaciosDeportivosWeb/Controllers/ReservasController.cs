using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReservasController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today);
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today.AddDays(6));
        if (hasta < desde) hasta = desde;

        var sedes = await spService.EspaciosComboSedesAsync(resolvedNegocioId.Value);
        if (!sedeId.HasValue && sedes.Count == 1 && int.TryParse(sedes[0].Value, out var sedeUnicaId))
            sedeId = sedeUnicaId;
        var espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value);
        var clientes = await spService.ReservasComboClientesAsync(resolvedNegocioId.Value);
        var reservas = await spService.ReservasListarAsync(resolvedNegocioId.Value, desde, hasta, sedeId, espacioDeportivoId, estado);

        var vm = new ReservasIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            FechaDesde = desde,
            FechaHasta = hasta,
            SedeId = sedeId,
            EspacioDeportivoId = espacioDeportivoId,
            Estado = estado,
            SedesFiltro = sedes,
            EspaciosFiltro = espacios,
            ClientesFiltro = clientes,
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
    public async Task<IActionResult> ObtenerReservaModal(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        var vm = await spService.ReservasObtenerAsync(negocioId, id);
        if (vm is null) return NotFound(new { ok = false, mensaje = "No se encontro la reserva." });

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
            adelanto = vm.Adelanto
        });
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
        if (!ModelState.IsValid)
        {
            var errores = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
            return BadRequest(new { ok = false, mensaje = "Revisa los datos ingresados.", errores });
        }
        if (EsFechaPasada(model.Fecha))
        {
            return BadRequest(new { ok = false, mensaje = "No se permite registrar reservas en fechas pasadas." });
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
            estadoCodigo = ObtenerCodigoEstado(r.Estado, r.TipoEvento),
            estadoTexto = ObtenerTextoEstado(r.Estado, r.TipoEvento),
            espacioDeportivoId = r.EspacioDeportivoId,
            backgroundColor = r.Color,
            borderColor = r.Color,
            editable = r.TipoEvento == "RESERVA" && baseVm.PuedeEditar
        });

        return Json(eventos);
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
    public async Task<IActionResult> CrearBloqueo(BloqueoHorarioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        if (!ModelState.IsValid)
        {
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }

        try
        {
            await spService.BloqueosCrearAsync(model, User.Identity?.Name ?? "sistema");
        }
        catch (Exception ex)
        {
            TempData["ReservasError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId, fechaDesde = model.Fecha, fechaHasta = model.Fecha });
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
        vm.Espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value);
        vm.Clientes = await spService.ReservasComboClientesAsync(resolvedNegocioId.Value);
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

        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId);
        model.Clientes = await spService.ReservasComboClientesAsync(model.NegocioId);
        if (EsFechaPasada(model.Fecha))
        {
            ModelState.AddModelError(nameof(model.Fecha), "No se permite registrar reservas en fechas pasadas.");
        }
        if (!ModelState.IsValid) return View(model);

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
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.Espacios = await spService.ReservasComboEspaciosAsync(resolvedNegocioId.Value);
        vm.Clientes = await spService.ReservasComboClientesAsync(resolvedNegocioId.Value);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ReservaFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId);
        model.Clientes = await spService.ReservasComboClientesAsync(model.NegocioId);
        if (EsFechaPasada(model.Fecha))
        {
            ModelState.AddModelError(nameof(model.Fecha), "No se permite registrar reservas en fechas pasadas.");
        }
        if (!ModelState.IsValid) return View(model);

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

        var ok = await spService.ReservasEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private static List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem> ObtenerEstadosFiltro()
    {
        return new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>
        {
            new("Todos", string.Empty),
            new("Pendiente", "1"),
            new("Confirmada", "2"),
            new("En uso", "3"),
            new("Finalizada", "4"),
            new("Cancelada", "5"),
            new("No asistio", "6")
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

    private static string? ObtenerCodigoEstado(int? estado, string? tipoEvento)
    {
        if (string.Equals(tipoEvento, "NO_ATENCION", StringComparison.OrdinalIgnoreCase))
            return "BLOQUEADO_NO_ATENCION";
        if (string.Equals(tipoEvento, "BLOQUEO", StringComparison.OrdinalIgnoreCase))
            return "BLOQUEADO";

        return estado switch
        {
            1 => "PENDIENTE",
            2 => "CONFIRMADA",
            3 => "EN_USO",
            4 => "FINALIZADA",
            5 => "CANCELADA",
            6 => "NO_SHOW",
            _ => null
        };
    }

    private static string? ObtenerTextoEstado(int? estado, string? tipoEvento)
    {
        if (string.Equals(tipoEvento, "NO_ATENCION", StringComparison.OrdinalIgnoreCase))
            return "Bloqueado/No atención";
        if (string.Equals(tipoEvento, "BLOQUEO", StringComparison.OrdinalIgnoreCase))
            return "Bloqueado";

        return estado switch
        {
            1 => "Pendiente",
            2 => "Confirmada",
            3 => "En uso",
            4 => "Finalizada",
            5 => "Cancelada",
            6 => "No show",
            _ => null
        };
    }
}
