using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text.Json;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class EspaciosController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService, ISedeImagenStorageService sedeImagenStorageService)
    : ModuloControllerBase(moduloPermisoService)
{
    private static readonly JsonSerializerOptions TarifaJsonSerializerOptions = new(JsonSerializerDefaults.Web);
    private const int MaxImagenesPorEspacio = 3;

    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new EspaciosIndexViewModel
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
            Espacios = await spService.EspaciosListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null))
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var configNegocio = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var espaciosActuales = await spService.EspaciosListarAsync(resolvedNegocioId.Value, null);
        var totalActivos = espaciosActuales.Count(x => string.Equals(x.Estado, "Activo", StringComparison.OrdinalIgnoreCase));
        var limiteEspacios = configNegocio?.EspaciosPermitidos ?? 6;
        if (totalActivos >= limiteEspacios)
        {
            TempData["EspaciosError"] = $"Limite de espacios alcanzado. Tu plan actual permite hasta {limiteEspacios} espacio(s) activos. Para continuar, solicita una ampliacion al administrador de plataforma.";
            return RedirectToAction(nameof(Index), new { negocioId = resolvedNegocioId.Value });
        }

        var vm = new EspacioFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual };
        await CargarCombosEspacioAsync(vm, AplicarSedeAsignada(baseVm, null));
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var configNegocio = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        var espaciosActuales = await spService.EspaciosListarAsync(model.NegocioId, null);
        var totalActivos = espaciosActuales.Count(x => string.Equals(x.Estado, "Activo", StringComparison.OrdinalIgnoreCase));
        var limiteEspacios = configNegocio?.EspaciosPermitidos ?? 6;
        if (totalActivos >= limiteEspacios)
            ModelState.AddModelError(string.Empty, $"Limite de espacios alcanzado. Tu plan actual permite hasta {limiteEspacios} espacio(s) activos. Para continuar, solicita una ampliacion al administrador de plataforma.");

        await CargarCombosEspacioAsync(model, AplicarSedeAsignada(baseVm, null));
        NormalizarEspaciosCompartidos(model);
        NormalizarFotosEspacio(model);
        AplicarEliminacionImagenes(model);
        ValidarCargaImagenesEspacio(model);
        if (!model.PuedeEditarTarifas)
            ModelState.AddModelError(string.Empty, "Debes configurar la moneda del club en Configuracion antes de registrar precios para espacios.");
        if (model.PuedeEditarTarifas)
        {
            CargarTarifasDesdeJson(model);
            ValidarTarifas(model);
            CargarTarifasFeriadoDesdeJson(model);
            ValidarTarifasFeriado(model);
        }
        else
        {
            model.Tarifas = new List<EspacioTarifaRangoViewModel>();
            model.TarifasJson = "[]";
            model.TarifasFeriado = new List<EspacioTarifaFeriadoRangoViewModel>();
            model.TarifasFeriadoJson = "[]";
        }
        if (!ModelState.IsValid) return View(model);

        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada.Value;

        var espacioId = await spService.EspaciosCrearAsync(model, User.Identity?.Name ?? "sistema");
        if (TieneImagenesNuevas(model))
        {
            model.Id = espacioId;
            await ProcesarCargaImagenesEspacioAsync(model);
            if (!ModelState.IsValid)
            {
                TempData["EspaciosError"] = ObtenerPrimerErrorModelState() ?? "El espacio se creo, pero no se pudieron subir sus imagenes. Completa la carga desde Editar espacio.";
                return RedirectToAction(nameof(Edit), new { id = espacioId, negocioId = model.NegocioId });
            }

            await spService.EspaciosActualizarAsync(model, User.Identity?.Name ?? "sistema");
        }
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.EspaciosObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!SedePermitida(baseVm, vm.SedeId))
            return Forbid();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCombosEspacioAsync(vm, AplicarSedeAsignada(baseVm, null));
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(EspacioFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue)
            model.SedeId = baseVm.SedeIdAsignada.Value;
        var urlsEliminar = ObtenerUrlsAEliminar(model);
        await CargarCombosEspacioAsync(model, AplicarSedeAsignada(baseVm, null));
        NormalizarEspaciosCompartidos(model);
        NormalizarFotosEspacio(model);
        AplicarEliminacionImagenes(model);
        await ProcesarCargaImagenesEspacioAsync(model);
        if (!model.PuedeEditarTarifas)
            ModelState.AddModelError(string.Empty, "Debes configurar la moneda del club en Configuracion antes de registrar precios para espacios.");
        if (model.PuedeEditarTarifas)
        {
            CargarTarifasDesdeJson(model);
            ValidarTarifas(model);
            CargarTarifasFeriadoDesdeJson(model);
            ValidarTarifasFeriado(model);
        }
        else
        {
            model.Tarifas = new List<EspacioTarifaRangoViewModel>();
            model.TarifasJson = "[]";
            model.TarifasFeriado = new List<EspacioTarifaFeriadoRangoViewModel>();
            model.TarifasFeriadoJson = "[]";
        }
        if (!ModelState.IsValid) return View(model);

        try
        {
            var ok = await spService.EspaciosActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar el espacio deportivo. Verifica el negocio seleccionado.");
                return View(model);
            }
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View(model);
        }

        if (urlsEliminar.Count > 0)
            await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsEliminar, HttpContext.RequestAborted);

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.EspaciosEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            if (!ok) return NotFound();
            TempData["EspaciosOk"] = "Espacio inactivado correctamente.";
        }
        catch (SqlException ex)
        {
            TempData["EspaciosError"] = ex.Message;
        }
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    private async Task CargarCombosEspacioAsync(EspacioFormViewModel model, int? sedeIdFiltro)
    {
        var tarifasJsonOriginal = model.TarifasJson;
        var tieneTarifasJson = !string.IsNullOrWhiteSpace(tarifasJsonOriginal);
        var tarifasFeriadoJsonOriginal = model.TarifasFeriadoJson;
        var tieneTarifasFeriadoJson = !string.IsNullOrWhiteSpace(tarifasFeriadoJsonOriginal);

        model.Sedes = await spService.EspaciosComboSedesAsync(model.NegocioId, sedeIdFiltro);
        model.TiposDeporte = await spService.EspaciosComboTiposDeporteAsync(model.NegocioId);
        model.TiposSuelo = await spService.EspaciosComboTiposSueloAsync(model.NegocioId);
        model.EspaciosCompartibles = model.SedeId > 0
            ? await spService.EspaciosComboCompartiblesAsync(model.NegocioId, model.SedeId, model.Id > 0 ? model.Id : null)
            : new List<SelectListItem>();
        InsertarOpcionSeleccione(model.Sedes, "Seleccione sede");
        InsertarOpcionSeleccione(model.TiposDeporte, "Seleccione deporte");
        InsertarOpcionSeleccione(model.TiposSuelo, "Seleccione tipo de suelo");
        await CargarMonedaConfiguradaAsync(model);
        model.TarifaDiasSemana =
        [
            new("Lunes", "1"),
            new("Martes", "2"),
            new("Miercoles", "3"),
            new("Jueves", "4"),
            new("Viernes", "5"),
            new("Sabado", "6"),
            new("Domingo", "0")
        ];

        if (!model.PuedeEditarTarifas)
        {
            model.Tarifas = new List<EspacioTarifaRangoViewModel>();
            model.TarifasJson = "[]";
            model.TarifasFeriado = new List<EspacioTarifaFeriadoRangoViewModel>();
            model.TarifasFeriadoJson = "[]";
            return;
        }

        if (model.Tarifas.Count == 0 && !tieneTarifasJson)
        {
            model.Tarifas =
            [
                new EspacioTarifaRangoViewModel
                {
                    DiaSemana = 1,
                    HoraInicio = new TimeOnly(8, 0),
                    HoraFin = new TimeOnly(9, 0),
                    Precio = 0.01m
                }
            ];
        }

        if (!tieneTarifasJson)
            model.TarifasJson = JsonSerializer.Serialize(model.Tarifas, TarifaJsonSerializerOptions);

        if (!tieneTarifasFeriadoJson)
            model.TarifasFeriadoJson = JsonSerializer.Serialize(model.TarifasFeriado, TarifaJsonSerializerOptions);
    }

    private static void InsertarOpcionSeleccione(List<SelectListItem> items, string texto)
    {
        if (items.Any(x => x.Value == "0"))
            return;

        items.Insert(0, new SelectListItem(texto, "0"));
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerEspaciosCompartibles(int negocioId, int sedeId, int? espacioActualId = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "ESPACIOS");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return Json(new { ok = false, mensaje = "No autorizado." });

        if (!baseVm.EsAdministrador && baseVm.SedeIdAsignada.HasValue && baseVm.SedeIdAsignada.Value != sedeId)
            return Json(new { ok = false, mensaje = "No autorizado para la sede seleccionada." });

        var items = await spService.EspaciosComboCompartiblesAsync(negocioId, sedeId, espacioActualId);
        return Json(new
        {
            ok = true,
            items = items.Select(x => new { value = x.Value, text = x.Text })
        });
    }

    [HttpGet]
    public async Task<IActionResult> VerImagen(string? url)
    {
        var imagen = await sedeImagenStorageService.ObtenerImagenVisualizacionAsync(url, HttpContext.RequestAborted);
        if (imagen is null)
            return NotFound();

        Response.Headers["Cache-Control"] = "public, max-age=300";
        return File(imagen.Value.Contenido, imagen.Value.ContentType);
    }

    private void NormalizarEspaciosCompartidos(EspacioFormViewModel model)
    {
        if (!model.TieneEspaciosCompartidos)
        {
            model.EspaciosBloqueoDirectoIds = new List<int>();
            model.EspaciosComponentesIds = new List<int>();
            return;
        }

        model.EspaciosBloqueoDirectoIds = model.EspaciosBloqueoDirectoIds
            .Where(x => x > 0 && x != model.Id)
            .Distinct()
            .ToList();

        model.EspaciosComponentesIds = model.EspaciosComponentesIds
            .Where(x => x > 0 && x != model.Id)
            .Distinct()
            .ToList();

        var idsRepetidos = model.EspaciosBloqueoDirectoIds
            .Intersect(model.EspaciosComponentesIds)
            .ToList();

        if (idsRepetidos.Count > 0)
            ModelState.AddModelError(nameof(model.EspaciosBloqueoDirectoIds), "Un mismo espacio no puede registrarse como bloqueo directo y como componente al mismo tiempo.");

        if (model.EspaciosBloqueoDirectoIds.Count == 0 && model.EspaciosComponentesIds.Count == 0)
            ModelState.AddModelError(nameof(model.EspaciosBloqueoDirectoIds), "Debes seleccionar al menos una relacion operativa.");
    }

    private async Task CargarMonedaConfiguradaAsync(EspacioFormViewModel model)
    {
        var configuracion = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        var monedas = await spService.ConfiguracionClubComboMonedasAsync(model.NegocioId);
        var monedaSeleccionada = configuracion is null
            ? null
            : monedas.FirstOrDefault(x => x.Value == configuracion.MonedaId.ToString());

        model.MonedaIdConfigurada = monedaSeleccionada is null ? null : configuracion!.MonedaId;
        model.MonedaEtiqueta = ResolverEtiquetaMoneda(monedaSeleccionada);
        model.PuedeEditarTarifas = model.MonedaIdConfigurada.HasValue && !string.IsNullOrWhiteSpace(model.MonedaEtiqueta);
    }

    private static string ResolverEtiquetaMoneda(SelectListItem? monedaSeleccionada)
    {
        if (monedaSeleccionada is null || string.IsNullOrWhiteSpace(monedaSeleccionada.Text))
            return string.Empty;

        var texto = monedaSeleccionada.Text;
        if (texto.Contains("(PEN)", StringComparison.OrdinalIgnoreCase))
            return "S/";
        if (texto.Contains("(USD)", StringComparison.OrdinalIgnoreCase))
            return "$";
        return texto;
    }

    private void CargarTarifasDesdeJson(EspacioFormViewModel model)
    {
        if (string.IsNullOrWhiteSpace(model.TarifasJson))
            return;

        try
        {
            var tarifas = JsonSerializer.Deserialize<List<EspacioTarifaRangoViewModel>>(model.TarifasJson, TarifaJsonSerializerOptions);
            model.Tarifas = tarifas ?? new List<EspacioTarifaRangoViewModel>();
        }
        catch
        {
            ModelState.AddModelError(string.Empty, "No se pudo leer el detalle de tarifas.");
        }
    }

    private void CargarTarifasFeriadoDesdeJson(EspacioFormViewModel model)
    {
        if (string.IsNullOrWhiteSpace(model.TarifasFeriadoJson))
            return;

        try
        {
            var tarifas = JsonSerializer.Deserialize<List<EspacioTarifaFeriadoRangoViewModel>>(model.TarifasFeriadoJson, TarifaJsonSerializerOptions);
            model.TarifasFeriado = tarifas ?? new List<EspacioTarifaFeriadoRangoViewModel>();
        }
        catch
        {
            ModelState.AddModelError(string.Empty, "No se pudo leer el detalle de tarifas por feriado.");
        }
    }

    private void ValidarTarifas(EspacioFormViewModel model)
    {
        if (model.Tarifas.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debes registrar al menos una tarifa.");
            return;
        }

        foreach (var tarifa in model.Tarifas)
        {
            if (tarifa.DiaSemana < 0 || tarifa.DiaSemana > 6)
                ModelState.AddModelError(string.Empty, "Existe una tarifa con dia invalido.");

            if (tarifa.HoraFin <= tarifa.HoraInicio)
                ModelState.AddModelError(string.Empty, "En tarifas, la hora fin debe ser mayor que la hora inicio.");

            if (tarifa.Precio <= 0)
                ModelState.AddModelError(string.Empty, "En tarifas, el precio debe ser mayor que cero.");
        }

        var tarifasPorDia = model.Tarifas.GroupBy(t => t.DiaSemana);
        foreach (var grupoDia in tarifasPorDia)
        {
            var ordenado = grupoDia.OrderBy(t => t.HoraInicio).ToList();
            for (var i = 1; i < ordenado.Count; i++)
            {
                if (ordenado[i].HoraInicio < ordenado[i - 1].HoraFin)
                {
                    ModelState.AddModelError(string.Empty, "Existen rangos de tarifa superpuestos en el mismo dia.");
                    break;
                }
            }
        }

        model.TarifasJson = JsonSerializer.Serialize(model.Tarifas, TarifaJsonSerializerOptions);
    }

    private void ValidarTarifasFeriado(EspacioFormViewModel model)
    {
        foreach (var tarifa in model.TarifasFeriado)
        {
            if (tarifa.HoraFin <= tarifa.HoraInicio)
                ModelState.AddModelError(string.Empty, "En tarifas por feriado, la hora fin debe ser mayor que la hora inicio.");

            if (tarifa.Precio <= 0)
                ModelState.AddModelError(string.Empty, "En tarifas por feriado, el precio debe ser mayor que cero.");
        }

        var ordenado = model.TarifasFeriado.OrderBy(t => t.HoraInicio).ToList();
        for (var i = 1; i < ordenado.Count; i++)
        {
            if (ordenado[i].HoraInicio < ordenado[i - 1].HoraFin)
            {
                ModelState.AddModelError(string.Empty, "Existen rangos de tarifas por feriado superpuestos.");
                break;
            }
        }

        model.TarifasFeriadoJson = JsonSerializer.Serialize(model.TarifasFeriado, TarifaJsonSerializerOptions);
    }

    private void NormalizarFotosEspacio(EspacioFormViewModel model)
    {
        model.FotoPrincipalUrl = string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) ? null : model.FotoPrincipalUrl.Trim();

        var fotos = (model.FotosUrlsCsv ?? string.Empty)
            .Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        model.FotosUrls = fotos;
        model.FotosUrlsCsv = fotos.Count == 0 ? null : string.Join(",", fotos);

        if (!string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && !Uri.IsWellFormedUriString(model.FotoPrincipalUrl, UriKind.Absolute))
            ModelState.AddModelError(nameof(model.FotoPrincipalUrl), "La foto principal debe ser una URL valida.");

        if (fotos.Any(url => !Uri.IsWellFormedUriString(url, UriKind.Absolute)))
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), "Todas las fotos de galeria deben ser URLs validas.");

        if (fotos.Count > MaxImagenesPorEspacio - 1)
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), $"La galeria permite maximo {MaxImagenesPorEspacio - 1} fotos alternativas.");

        var totalFotos = (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) ? 0 : 1) + fotos.Count;
        if (totalFotos > MaxImagenesPorEspacio)
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), $"Solo se permiten {MaxImagenesPorEspacio} imagenes por espacio deportivo.");

        if (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && fotos.Count > 0)
            ModelState.AddModelError(nameof(model.FotoPrincipalUrl), "Debes tener una foto principal cuando registres fotos alternativas.");
    }

    private static void AplicarEliminacionImagenes(EspacioFormViewModel model)
    {
        var aEliminar = (model.FotosEliminarUrls ?? new List<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (aEliminar.Count == 0)
            return;

        if (!string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && aEliminar.Contains(model.FotoPrincipalUrl))
            model.FotoPrincipalUrl = null;

        model.FotosUrls = (model.FotosUrls ?? new List<string>())
            .Where(x => !aEliminar.Contains(x))
            .ToList();

        if (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && model.FotosUrls.Count > 0)
        {
            model.FotoPrincipalUrl = model.FotosUrls[0];
            model.FotosUrls.RemoveAt(0);
        }

        model.FotosUrlsCsv = model.FotosUrls.Count == 0 ? null : string.Join(",", model.FotosUrls);
    }

    private static List<string> ObtenerUrlsAEliminar(EspacioFormViewModel model)
    {
        return (model.FotosEliminarUrls ?? new List<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void ValidarCargaImagenesEspacio(EspacioFormViewModel model)
    {
        var archivos = (model.ImagenesArchivos ?? new List<IFormFile>())
            .Where(f => f is not null && f.Length > 0)
            .ToList();

        if (archivos.Count == 0)
            return;

        if (archivos.Count > MaxImagenesPorEspacio)
        {
            ModelState.AddModelError(nameof(model.ImagenesArchivos), $"Solo se permiten {MaxImagenesPorEspacio} imagenes por espacio deportivo.");
            return;
        }

        var totalActual = (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) ? 0 : 1) + (model.FotosUrls?.Count ?? 0);
        if (totalActual + archivos.Count > MaxImagenesPorEspacio)
            ModelState.AddModelError(nameof(model.ImagenesArchivos), $"Solo se permiten {MaxImagenesPorEspacio} imagenes por espacio deportivo. Ya tienes {totalActual} registradas.");
    }

    private async Task ProcesarCargaImagenesEspacioAsync(EspacioFormViewModel model)
    {
        var archivos = (model.ImagenesArchivos ?? new List<IFormFile>())
            .Where(f => f is not null && f.Length > 0)
            .ToList();

        if (archivos.Count == 0)
            return;

        if (model.Id <= 0)
        {
            ModelState.AddModelError(nameof(model.ImagenesArchivos), "Primero se debe guardar el espacio para poder asociar sus imagenes.");
            return;
        }

        if (!ModelState.IsValid)
            return;

        try
        {
            var urls = await sedeImagenStorageService.UploadEspacioImagenesAsync(
                model.NegocioId,
                model.Id,
                archivos,
                HttpContext.RequestAborted);

            if (urls.Count == 0)
            {
                ModelState.AddModelError(nameof(model.ImagenesArchivos), "No se pudo completar la carga de imagenes.");
                return;
            }

            var urlsActuales = new List<string>();
            if (!string.IsNullOrWhiteSpace(model.FotoPrincipalUrl))
                urlsActuales.Add(model.FotoPrincipalUrl);
            if (model.FotosUrls?.Count > 0)
                urlsActuales.AddRange(model.FotosUrls.Where(x => !string.IsNullOrWhiteSpace(x)));

            var totalFinal = urlsActuales.Count + urls.Count;
            if (totalFinal > MaxImagenesPorEspacio)
            {
                ModelState.AddModelError(nameof(model.ImagenesArchivos), $"Solo se permiten {MaxImagenesPorEspacio} imagenes por espacio deportivo. Ya tienes {urlsActuales.Count} registradas.");
                return;
            }

            urlsActuales.AddRange(urls);
            model.FotoPrincipalUrl = urlsActuales[0];
            model.FotosUrls = urlsActuales.Skip(1).ToList();
            model.FotosUrlsCsv = model.FotosUrls.Count == 0 ? null : string.Join(",", model.FotosUrls);
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(nameof(model.ImagenesArchivos), $"No se pudieron subir las imagenes: {ex.Message}");
        }
    }

    private static bool TieneImagenesNuevas(EspacioFormViewModel model)
    {
        return (model.ImagenesArchivos ?? new List<IFormFile>()).Any(x => x is not null && x.Length > 0);
    }

    private string? ObtenerPrimerErrorModelState()
    {
        return ModelState.Values
            .SelectMany(x => x.Errors)
            .Select(x => x.ErrorMessage)
            .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
    }
}
