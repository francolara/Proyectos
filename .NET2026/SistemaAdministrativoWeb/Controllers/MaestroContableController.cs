using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;
using SistemaAdministrativoWeb.ViewModels.Plataforma;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize(Roles = "SuperAdmin")]
public sealed class MaestroContableController(IMaestroContableRepository repository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoPaginaAyuda = 40;

    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        PrepararShell();
        var cuentas = await repository.ListarPlanCuentaAsync(null, null, false, false, 1, 1, cancellationToken);
        var destinos = await repository.ListarCuentasDestinoAsync(null, 1, 1, cancellationToken);
        var origenes = await repository.ListarOrigenesAsync(null, false, 1, 1, cancellationToken);
        var configuraciones = await repository.ListarConfiguracionesAsync(null, 1, 1, cancellationToken);
        var incidencias = await repository.ValidarAsync(cancellationToken);
        return View(new MaestroContableIndexViewModel
        {
            TotalCuentas = cuentas.TotalRecords,
            TotalReglasDestino = destinos.TotalRecords,
            TotalOrigenes = origenes.TotalRecords,
            TotalConfiguraciones = configuraciones.TotalRecords,
            Incidencias = incidencias
        });
    }

    [HttpGet]
    public async Task<IActionResult> PlanCuenta(string? textoBusqueda, byte? nivelCuenta, int pagina = 1, CancellationToken cancellationToken = default)
    {
        PrepararShell(); pagina = Math.Max(1, pagina);
        var resultado = await repository.ListarPlanCuentaAsync(textoBusqueda, nivelCuenta, false, false, pagina, TamanoPagina, cancellationToken);
        return View(new PlanCuentaMaestroIndexViewModel
        {
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty, NivelCuenta = nivelCuenta, Items = resultado.Items,
            Paginacion = CrearPaginacion(resultado)
        });
    }

    [HttpGet]
    public IActionResult RegistrarPlanCuenta() { PrepararShell(); return View("PlanCuentaFormulario", new PlanCuentaMaestroFormViewModel()); }

    [HttpGet]
    public async Task<IActionResult> EditarPlanCuenta(int id, CancellationToken cancellationToken)
    {
        PrepararShell(); var item = await repository.ObtenerPlanCuentaAsync(id, cancellationToken); if (item is null) return NotFound();
        return View("PlanCuentaFormulario", new PlanCuentaMaestroFormViewModel
        {
            IdPlanCuentaMaestro = item.IdPlanCuentaMaestro, CodigoCuenta = item.CodigoCuenta, CodigoCuentaPadre = item.CodigoCuentaPadre,
            NombreCuenta = item.NombreCuenta, ColBalance = item.ColBalance, IdMoneda = item.IdMoneda, TipoCambio = item.TipoCambio,
            AceptaMovimiento = item.AceptaMovimiento, RequiereCentroCosto = item.RequiereCentroCosto, Estado = item.Estado, Orden = item.Orden
        });
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarPlanCuenta(PlanCuentaMaestroFormViewModel model, CancellationToken cancellationToken)
    {
        PrepararShell(); if (!ModelState.IsValid) return View("PlanCuentaFormulario", model);
        try
        {
            await repository.GuardarPlanCuentaAsync(new GuardarPlanCuentaMaestroRequest
            {
                IdPlanCuentaMaestro = model.IdPlanCuentaMaestro, CodigoCuenta = model.CodigoCuenta.Trim(), CodigoCuentaPadre = Limpiar(model.CodigoCuentaPadre),
                NombreCuenta = model.NombreCuenta.Trim(), ColBalance = model.ColBalance.Trim().ToUpperInvariant(), IdMoneda = (model.IdMoneda ?? string.Empty).Trim().ToUpperInvariant(),
                TipoCambio = (model.TipoCambio ?? string.Empty).Trim().ToUpperInvariant(), AceptaMovimiento = model.AceptaMovimiento, RequiereCentroCosto = model.RequiereCentroCosto,
                Estado = model.Estado, Orden = model.Orden, UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);
            TempData["MaestroOk"] = model.IdPlanCuentaMaestro.HasValue ? "Cuenta maestra actualizada." : "Cuenta maestra creada.";
            return RedirectToAction(nameof(PlanCuenta));
        }
        catch (Exception ex) { ModelState.AddModelError(string.Empty, ex.Message); return View("PlanCuentaFormulario", model); }
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarPlanCuenta(int id, CancellationToken cancellationToken) => await EjecutarYVolverAsync(() => repository.EliminarPlanCuentaAsync(id, cancellationToken), nameof(PlanCuenta), "Cuenta maestra eliminada.");

    [HttpGet]
    public async Task<IActionResult> BuscarCuentas(string? texto, bool soloMovimiento = true, int pagina = 1, CancellationToken cancellationToken = default)
    {
        var resultado = await repository.ListarPlanCuentaAsync(texto, null, soloMovimiento, true, Math.Max(1, pagina), TamanoPaginaAyuda, cancellationToken);
        return Json(new { items = resultado.Items.Select(x => new { codigo = x.CodigoCuenta, nombre = x.NombreCuenta, nivel = x.NivelCuenta }), pagina = resultado.PageNumber, totalPaginas = Math.Max(1, (int)Math.Ceiling(resultado.TotalRecords / (double)TamanoPaginaAyuda)) });
    }

    [HttpGet]
    public async Task<IActionResult> CuentasDestino(string? textoBusqueda, int pagina = 1, CancellationToken cancellationToken = default)
    {
        PrepararShell(); pagina = Math.Max(1, pagina); var resultado = await repository.ListarCuentasDestinoAsync(textoBusqueda, pagina, TamanoPagina, cancellationToken);
        return View(new CuentaDestinoMaestroIndexViewModel { TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty, Items = resultado.Items, Paginacion = CrearPaginacion(resultado) });
    }

    [HttpGet]
    public IActionResult RegistrarCuentaDestino() { PrepararShell(); return View("CuentaDestinoFormulario", new CuentaDestinoMaestroFormViewModel()); }

    [HttpGet]
    public async Task<IActionResult> EditarCuentaDestino(int id, CancellationToken cancellationToken)
    {
        PrepararShell(); var item = await repository.ObtenerCuentaDestinoAsync(id, cancellationToken); if (item is null) return NotFound();
        return View("CuentaDestinoFormulario", new CuentaDestinoMaestroFormViewModel
        {
            IdCuentaDestinoReglaMaestro = item.IdCuentaDestinoReglaMaestro, CodigoCuentaOrigen = item.CodigoCuentaOrigen, Activo = item.Activo, Observacion = item.Observacion,
            Detalles = item.Detalles.Select(x => new CuentaDestinoDetalleMaestroFormViewModel { Orden = x.Orden, CodigoCuentaDestinoCargo = x.CodigoCuentaDestinoCargo, CodigoCuentaDestinoAbono = x.CodigoCuentaDestinoAbono, Porcentaje = x.Porcentaje, Activo = x.Activo }).ToList()
        });
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarCuentaDestino(CuentaDestinoMaestroFormViewModel model, CancellationToken cancellationToken)
    {
        PrepararShell();
        if (model.Detalles.Count == 0) ModelState.AddModelError(nameof(model.Detalles), "Debe registrar al menos un tramo.");
        if (!ModelState.IsValid) return View("CuentaDestinoFormulario", model);
        try
        {
            await repository.GuardarCuentaDestinoAsync(new GuardarCuentaDestinoMaestroRequest
            {
                IdCuentaDestinoReglaMaestro = model.IdCuentaDestinoReglaMaestro, CodigoCuentaOrigen = model.CodigoCuentaOrigen.Trim(), Activo = model.Activo,
                Observacion = Limpiar(model.Observacion), UsuarioRegistro = User.Identity?.Name,
                Detalles = model.Detalles.Select(x => new GuardarCuentaDestinoDetalleMaestroRequest { Orden = x.Orden, CodigoCuentaDestinoCargo = x.CodigoCuentaDestinoCargo.Trim(), CodigoCuentaDestinoAbono = x.CodigoCuentaDestinoAbono.Trim(), Porcentaje = x.Porcentaje, Activo = x.Activo }).ToList()
            }, cancellationToken);
            TempData["MaestroOk"] = model.IdCuentaDestinoReglaMaestro.HasValue ? "Regla maestra actualizada." : "Regla maestra creada.";
            return RedirectToAction(nameof(CuentasDestino));
        }
        catch (Exception ex) { ModelState.AddModelError(string.Empty, ex.Message); return View("CuentaDestinoFormulario", model); }
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarCuentaDestino(int id, CancellationToken cancellationToken) => await EjecutarYVolverAsync(() => repository.EliminarCuentaDestinoAsync(id, cancellationToken), nameof(CuentasDestino), "Regla maestra eliminada.");

    [HttpGet]
    public async Task<IActionResult> Asignaciones(CancellationToken cancellationToken)
    {
        PrepararShell(); var resultado = await repository.ListarAsignacionesAsync(cancellationToken);
        return View(new AsignacionesMaestroViewModel { Parametros = resultado.Parametros, Impuestos = resultado.Impuestos, Documentos = resultado.Documentos });
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarAsignacion(AsignacionCuentaMaestroFormViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid) { TempData["MaestroError"] = "La asignación enviada no es válida."; return RedirectToAction(nameof(Asignaciones)); }
        try
        {
            await repository.GuardarAsignacionAsync(new GuardarAsignacionMaestroRequest
            {
                TipoAsignacion = model.TipoAsignacion, IdRegistro = model.IdRegistro, CodigoCuenta = Limpiar(model.CodigoCuenta),
                CodigoCuentaVentaSoles = Limpiar(model.CodigoCuentaVentaSoles), CodigoCuentaVentaDolares = Limpiar(model.CodigoCuentaVentaDolares),
                CodigoCuentaCompraSoles = Limpiar(model.CodigoCuentaCompraSoles), CodigoCuentaCompraDolares = Limpiar(model.CodigoCuentaCompraDolares), UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);
            TempData["MaestroOk"] = "Asignación contable actualizada.";
        }
        catch (Exception ex) { TempData["MaestroError"] = ex.Message; }
        return RedirectToAction(nameof(Asignaciones));
    }

    [HttpGet]
    public async Task<IActionResult> Origenes(string? textoBusqueda, int pagina = 1, CancellationToken cancellationToken = default)
    {
        PrepararShell(); pagina = Math.Max(1, pagina); var resultado = await repository.ListarOrigenesAsync(textoBusqueda, false, pagina, TamanoPagina, cancellationToken);
        return View(new OrigenMaestroIndexViewModel { TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty, Items = resultado.Items, Paginacion = CrearPaginacion(resultado) });
    }

    [HttpGet]
    public IActionResult RegistrarOrigen() { PrepararShell(); return View("OrigenFormulario", new OrigenMaestroFormViewModel()); }

    [HttpGet]
    public async Task<IActionResult> EditarOrigen(int id, CancellationToken cancellationToken)
    {
        PrepararShell(); var item = await repository.ObtenerOrigenAsync(id, cancellationToken); if (item is null) return NotFound();
        return View("OrigenFormulario", new OrigenMaestroFormViewModel { IdOrigenMaestro = item.IdOrigenMaestro, CodigoOrigen = item.CodigoOrigen, NombreOrigen = item.NombreOrigen, ModuloOrigen = item.ModuloOrigen, PermiteRegistroManual = item.PermiteRegistroManual, Estado = item.Estado, Orden = item.Orden });
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarOrigen(OrigenMaestroFormViewModel model, CancellationToken cancellationToken)
    {
        PrepararShell(); if (!ModelState.IsValid) return View("OrigenFormulario", model);
        try
        {
            await repository.GuardarOrigenAsync(new GuardarOrigenMaestroRequest { IdOrigenMaestro = model.IdOrigenMaestro, CodigoOrigen = model.CodigoOrigen.Trim().ToUpperInvariant(), NombreOrigen = model.NombreOrigen.Trim(), ModuloOrigen = model.ModuloOrigen.Trim().ToUpperInvariant(), PermiteRegistroManual = model.PermiteRegistroManual, Estado = model.Estado, Orden = model.Orden, UsuarioRegistro = User.Identity?.Name }, cancellationToken);
            TempData["MaestroOk"] = model.IdOrigenMaestro.HasValue ? "Origen maestro actualizado." : "Origen maestro creado."; return RedirectToAction(nameof(Origenes));
        }
        catch (Exception ex) { ModelState.AddModelError(string.Empty, ex.Message); return View("OrigenFormulario", model); }
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarOrigen(int id, CancellationToken cancellationToken) => await EjecutarYVolverAsync(() => repository.EliminarOrigenAsync(id, cancellationToken), nameof(Origenes), "Origen maestro eliminado.");

    [HttpGet]
    public async Task<IActionResult> Configuracion(string? textoBusqueda, int pagina = 1, CancellationToken cancellationToken = default)
    {
        PrepararShell(); pagina = Math.Max(1, pagina); var resultado = await repository.ListarConfiguracionesAsync(textoBusqueda, pagina, TamanoPagina, cancellationToken);
        var origenes = await repository.ListarOrigenesAsync(null, true, 1, 500, cancellationToken);
        return View(new ConfiguracionMaestroIndexViewModel { TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty, Items = resultado.Items, OrigenesActivos = origenes.Items, Paginacion = CrearPaginacion(resultado) });
    }

    [HttpPost, ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarConfiguracion(ConfiguracionOrigenMaestroFormViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid) { TempData["MaestroError"] = "Debe seleccionar un origen válido."; return RedirectToAction(nameof(Configuracion)); }
        try { await repository.GuardarOrigenConfiguracionAsync(model.IdConfiguracionContabilizacionMaestro, model.CodigoOrigen, User.Identity?.Name, cancellationToken); TempData["MaestroOk"] = "Origen de la configuración actualizado."; }
        catch (Exception ex) { TempData["MaestroError"] = ex.Message; }
        return RedirectToAction(nameof(Configuracion));
    }

    private async Task<IActionResult> EjecutarYVolverAsync(Func<Task> accion, string destino, string mensaje)
    {
        try { await accion(); TempData["MaestroOk"] = mensaje; } catch (Exception ex) { TempData["MaestroError"] = ex.Message; }
        return RedirectToAction(destino);
    }

    private void PrepararShell() => ViewData["PlatformShell"] = true;
    private static string? Limpiar(string? valor) => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();
    private static PaginacionViewModel CrearPaginacion<T>(PagedResult<T> resultado) => new() { PaginaActual = resultado.PageNumber, TamanoPagina = resultado.PageSize, TotalRegistros = resultado.TotalRecords };
}
