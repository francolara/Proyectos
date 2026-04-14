using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class PromocionesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
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

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var (desde, hasta) = ResolverRangoFechas(fechaDesde, fechaHasta, preset);
        var estadoFiltro = NormalizarEstadoFiltro(estado);
        bool? soloActivos = estadoFiltro switch
        {
            "inactivos" => false,
            "todos" => null,
            _ => true
        };
        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var (promociones, totalRegistros) = await spService.PromocionesListarAsync(
            resolvedNegocioId.Value,
            AplicarSedeAsignada(baseVm, null),
            desde,
            hasta,
            soloActivos,
            paginaActual,
            tamanoPagina);
        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
        {
            paginaActual = totalPaginas;
            (promociones, totalRegistros) = await spService.PromocionesListarAsync(
                resolvedNegocioId.Value,
                AplicarSedeAsignada(baseVm, null),
                desde,
                hasta,
                soloActivos,
                paginaActual,
                tamanoPagina);
        }

        var vm = new PromocionesIndexViewModel
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
            Promociones = promociones
        };

        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new PromocionFormViewModel
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
    public async Task<IActionResult> Create(PromocionFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }
        if (model.PorcentajeDescuento != decimal.Truncate(model.PorcentajeDescuento))
        {
            ModelState.AddModelError(nameof(model.PorcentajeDescuento), "El porcentaje solo permite numeros enteros.");
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        await spService.PromocionesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.PromocionesObtenerAsync(resolvedNegocioId.Value, id);
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
    public async Task<IActionResult> Edit(PromocionFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada;

        if (!ModelState.IsValid)
        {
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }
        if (model.PorcentajeDescuento != decimal.Truncate(model.PorcentajeDescuento))
        {
            ModelState.AddModelError(nameof(model.PorcentajeDescuento), "El porcentaje solo permite numeros enteros.");
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }

        var ok = await spService.PromocionesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo guardar la promocion. Verifica el negocio seleccionado.");
            await CargarCombosAsync(model, baseVm);
            return View(model);
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "PROMOCIONES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.PromocionesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private async Task CargarCombosAsync(PromocionFormViewModel model, ModuloBaseViewModel baseVm)
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

    private static string NormalizarEstadoFiltro(string? estado)
    {
        var filtro = (estado ?? string.Empty).Trim().ToLowerInvariant();
        return filtro switch
        {
            "todos" => "todos",
            "inactivos" => "inactivos",
            _ => "activos"
        };
    }
}
