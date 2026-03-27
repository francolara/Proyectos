using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReservasController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, int? espacioDeportivoId, int? estado)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today);
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today.AddDays(6));
        if (hasta < desde) hasta = desde;

        var sedes = await spService.EspaciosComboSedesAsync(negocioId);
        var espacios = await spService.ReservasComboEspaciosAsync(negocioId);
        var reservas = await spService.ReservasListarAsync(negocioId, desde, hasta, sedeId, espacioDeportivoId, estado);

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
            EstadosFiltro = ObtenerEstadosFiltro(),
            Reservas = reservas,
            Bloqueos = await spService.BloqueosListarAsync(negocioId, desde, hasta, sedeId, espacioDeportivoId),
            BloqueoForm = new BloqueoHorarioFormViewModel
            {
                NegocioId = negocioId,
                Fecha = desde,
                Espacios = espacios
            }
        };
        return View(vm);
    }

    [HttpGet]
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

        if (request.Inicio >= request.Fin)
        {
            return BadRequest(new { ok = false, mensaje = "El horario no es valido." });
        }

        try
        {
            var ok = await spService.ReservasMoverAsync(
                request.NegocioId,
                request.ReservaId,
                DateOnly.FromDateTime(request.Inicio),
                TimeOnly.FromDateTime(request.Inicio),
                TimeOnly.FromDateTime(request.Fin),
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

    public async Task<IActionResult> Create(int negocioId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new ReservaFormViewModel { NegocioId = negocioId, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        vm.Espacios = await spService.ReservasComboEspaciosAsync(negocioId);
        vm.Clientes = await spService.ReservasComboClientesAsync(negocioId);
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

    public async Task<IActionResult> Edit(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "RESERVAS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ReservasObtenerAsync(negocioId, id);
        if (vm is null) return NotFound();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        vm.Espacios = await spService.ReservasComboEspaciosAsync(negocioId);
        vm.Clientes = await spService.ReservasComboClientesAsync(negocioId);
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
        if (!ModelState.IsValid) return View(model);

        try
        {
            var ok = await spService.ReservasActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok) return NotFound();
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
}
