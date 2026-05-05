using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class CuponesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(
        int? negocioId,
        DateOnly? fechaDesde = null,
        DateOnly? fechaHasta = null,
        string? estado = null,
        string? preset = null,
        int pagina = 1)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CUPONES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var (desde, hasta) = ResolverRangoFechas(fechaDesde, fechaHasta, preset);
        if (hasta < desde) (desde, hasta) = (hasta, desde);

        var estadoFiltro = (estado ?? "vigentes").Trim().ToLowerInvariant();
        if (estadoFiltro is not ("vigentes" or "activos" or "agotados" or "vencidos" or "todos"))
            estadoFiltro = "vigentes";

        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var (cupones, totalRegistros) = await spService.CuponesListarAsync(
            resolvedNegocioId.Value,
            AplicarSedeAsignada(baseVm, null),
            desde,
            hasta,
            estadoFiltro,
            paginaActual,
            tamanoPagina);

        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
        {
            paginaActual = totalPaginas;
            (cupones, totalRegistros) = await spService.CuponesListarAsync(
                resolvedNegocioId.Value,
                AplicarSedeAsignada(baseVm, null),
                desde,
                hasta,
                estadoFiltro,
                paginaActual,
                tamanoPagina);
        }

        var vm = new CuponesIndexViewModel
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
            EstadoFiltro = estadoFiltro,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            Cupones = cupones
        };

        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CUPONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new CuponFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            Activo = true
        };

        await CargarCombosAsync(vm, baseVm);
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            vm.SedeId = baseVm.SedeIdAsignada;

        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(CuponFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CUPONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        await spService.CuponesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CUPONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.CuponesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!SedePermitida(baseVm, vm.SedeId))
            return Forbid();

        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCombosAsync(vm, baseVm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(CuponFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CUPONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        var ok = await spService.CuponesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar el cupón.");
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "CUPONES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.CuponesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();

        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [HttpGet]
    public async Task<IActionResult> EspaciosPorSede(int negocioId, int? sedeId)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "CUPONES");
        if (baseVm is null || (!baseVm.PuedeCrear && !baseVm.PuedeEditar))
            return Forbid();

        var sedeFiltro = AplicarSedeAsignada(baseVm, sedeId);
        var espacios = await spService.ReservasComboEspaciosAsync(negocioId, sedeFiltro);
        var payload = espacios.Select(x => new { value = x.Value, text = x.Text }).ToList();
        return Json(payload);
    }

    private async Task CargarCombosAsync(CuponFormViewModel model, ModuloBaseViewModel baseVm)
    {
        var sedeFiltro = AplicarSedeAsignada(baseVm, null);
        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId, sedeFiltro);
        model.Espacios = await spService.ReservasComboEspaciosAsync(model.NegocioId, sedeFiltro);

        if (baseVm.EsAdministrador)
            model.Sedes.Insert(0, new SelectListItem("Todas las sedes", string.Empty));
        model.Espacios.Insert(0, new SelectListItem("Todos los espacios", string.Empty));
    }

    private static (DateOnly Desde, DateOnly Hasta) ResolverRangoFechas(DateOnly? fechaDesde, DateOnly? fechaHasta, string? preset)
    {
        var hoy = DateOnly.FromDateTime(DateTime.Today);
        return (preset ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "hoy" => (hoy, hoy),
            "7d" => (hoy.AddDays(-6), hoy),
            "30d" => (hoy.AddDays(-29), hoy),
            "mes" => (new DateOnly(hoy.Year, hoy.Month, 1), new DateOnly(hoy.Year, hoy.Month, DateTime.DaysInMonth(hoy.Year, hoy.Month))),
            _ => (fechaDesde ?? hoy.AddDays(-6), fechaHasta ?? hoy.AddDays(30))
        };
    }
}
