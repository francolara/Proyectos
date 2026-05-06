using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class SedesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService, ISedeImagenStorageService sedeImagenStorageService)
    : ModuloControllerBase(moduloPermisoService)
{
    private const int MaxImagenesPorSede = 6;

    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = new SedesIndexViewModel
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
            Sedes = await spService.SedesListarAsync(resolvedNegocioId.Value, AplicarSedeAsignada(baseVm, null))
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var configNegocio = await spService.ConfiguracionClubObtenerAsync(resolvedNegocioId.Value);
        var sedesActuales = await spService.SedesListarAsync(resolvedNegocioId.Value, null);
        var totalActivas = sedesActuales.Count(x => x.Activo);
        var limiteSedes = configNegocio?.SedesPermitidas ?? 2;
        if (totalActivas >= limiteSedes)
        {
            TempData["SedesError"] = $"Limite de sedes alcanzado. Tu plan actual permite hasta {limiteSedes} sede(s) activas. Para continuar, solicita una ampliacion al administrador de plataforma.";
            return RedirectToAction(nameof(Index), new { negocioId = resolvedNegocioId.Value });
        }

        var vm = new SedeFormViewModel { NegocioId = resolvedNegocioId.Value, NegocioNombre = baseVm.NegocioNombre, RolActual = baseVm.RolActual, Activo = true };
        await CargarCatalogoSedeAsync(vm);
        InicializarTelefonosParaVista(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(SedeFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var configNegocio = await spService.ConfiguracionClubObtenerAsync(model.NegocioId);
        var sedesActuales = await spService.SedesListarAsync(model.NegocioId, null);
        var totalActivas = sedesActuales.Count(x => x.Activo);
        var limiteSedes = configNegocio?.SedesPermitidas ?? 2;
        if (totalActivas >= limiteSedes)
            ModelState.AddModelError(string.Empty, $"Limite de sedes alcanzado. Tu plan actual permite hasta {limiteSedes} sede(s) activas. Para continuar, solicita una ampliacion al administrador de plataforma.");

        ComponerTelefonos(model);
        NormalizarUbicacionYFotos(model);
        await ValidarUbigeoSedeAsync(model);
        AplicarEliminacionImagenes(model);
        await ProcesarCargaImagenesAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCatalogoSedeAsync(model);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        await spService.SedesCrearAsync(model, User.Identity?.Name ?? "sistema");
        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.SedesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        if (!SedePermitida(baseVm, vm.Id))
            return Forbid();
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        await CargarCatalogoSedeAsync(vm);
        vm.SeriesDocumentoConfig = await spService.SedesSeriesDocumentoListarAsync(resolvedNegocioId.Value, vm.Id);
        InicializarTelefonosParaVista(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(SedeFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        if (!SedePermitida(baseVm, model.Id))
            return Forbid();
        var urlsEliminar = ObtenerUrlsAEliminar(model);
        ComponerTelefonos(model);
        NormalizarUbicacionYFotos(model);
        await ValidarUbigeoSedeAsync(model);
        AplicarEliminacionImagenes(model);
        await ProcesarCargaImagenesAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCatalogoSedeAsync(model);
            model.SeriesDocumentoConfig = await spService.SedesSeriesDocumentoListarAsync(model.NegocioId, model.Id);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        var ok = await spService.SedesActualizarAsync(model, User.Identity?.Name ?? "sistema");
        if (!ok)
        {
            ModelState.AddModelError(string.Empty, "No se pudo actualizar la sede. Verifica el negocio seleccionado.");
            await CargarCatalogoSedeAsync(model);
            model.SeriesDocumentoConfig = await spService.SedesSeriesDocumentoListarAsync(model.NegocioId, model.Id);
            model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
            return View(model);
        }

        foreach (var item in model.SeriesDocumentoConfig ?? new List<SedeSerieDocumentoConfigItemViewModel>())
        {
            if (item.PermiteMultiplesSeries)
            {
                await spService.SedesSeriesDocumentoGuardarMultiplesAsync(
                    model.NegocioId,
                    model.Id,
                    item.CodigoSunat,
                    item.NegocioSeriesIds,
                    User.Identity?.Name ?? "sistema");
            }
            else
            {
                await spService.SedesSeriesDocumentoGuardarAsync(
                    model.NegocioId,
                    model.Id,
                    item.CodigoSunat,
                    item.NegocioSerieId,
                    User.Identity?.Name ?? "sistema");
            }
        }

        if (urlsEliminar.Count > 0)
            await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsEliminar, HttpContext.RequestAborted);

        return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "SEDES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var ok = await spService.SedesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
        if (!ok) return NotFound();
        return RedirectToAction(nameof(Index), new { negocioId });
    }

    [AllowAnonymous]
    [HttpGet]
    public async Task<IActionResult> VerImagen(string? url)
    {
        var imagen = await sedeImagenStorageService.ObtenerImagenVisualizacionAsync(url, HttpContext.RequestAborted);
        if (imagen is null)
            return NotFound();

        Response.Headers["Cache-Control"] = "public, max-age=300";
        return File(imagen.Value.Contenido, imagen.Value.ContentType);
    }

    private async Task CargarCatalogoSedeAsync(SedeFormViewModel model)
    {
        model.ServiciosDisponibles = await spService.SedesComboServiciosAsync();
        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();
        model.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(model.CodigoDepartamento)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
        model.DistritosUbigeo = !string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(model.CodigoProvincia)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
    }

    private async Task ValidarUbigeoSedeAsync(SedeFormViewModel model)
    {
        model.CodigoDepartamento = string.IsNullOrWhiteSpace(model.CodigoDepartamento) ? null : model.CodigoDepartamento.Trim();
        model.CodigoProvincia = string.IsNullOrWhiteSpace(model.CodigoProvincia) ? null : model.CodigoProvincia.Trim();
        model.CodigoUbigeo = string.IsNullOrWhiteSpace(model.CodigoUbigeo) ? string.Empty : model.CodigoUbigeo.Trim();

        if (model.CodigoUbigeo.Length != 6)
        {
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Debes seleccionar un distrito valido.");
            return;
        }

        var ubigeo = await spService.UbigeoObtenerPorCodigoAsync(model.CodigoUbigeo);
        if (ubigeo is null)
        {
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Debes seleccionar un distrito valido.");
            return;
        }

        model.CodigoDepartamento = ubigeo.CodigoDepartamento;
        model.CodigoProvincia = ubigeo.CodigoProvincia;
    }

    private static void InicializarTelefonosParaVista(SedeFormViewModel model)
    {
        TelefonoInternacionalHelper.Descomponer(model.Telefono, out var telefonoCodigoPais, out var telefonoNumeroLocal);
        model.TelefonoCodigoPais = telefonoCodigoPais;
        model.TelefonoNumeroLocal = telefonoNumeroLocal;

        TelefonoInternacionalHelper.Descomponer(model.WhatsappContacto, out var whatsappCodigoPais, out var whatsappNumeroLocal);
        model.WhatsappCodigoPais = whatsappCodigoPais;
        model.WhatsappNumeroLocal = whatsappNumeroLocal;
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
    }

    private static void ComponerTelefonos(SedeFormViewModel model)
    {
        model.Telefono = TelefonoInternacionalHelper.Componer(model.TelefonoCodigoPais, model.TelefonoNumeroLocal);
        model.WhatsappContacto = TelefonoInternacionalHelper.Componer(model.WhatsappCodigoPais, model.WhatsappNumeroLocal);
    }

    private void NormalizarUbicacionYFotos(SedeFormViewModel model)
    {
        model.ConsideracionesReserva = string.IsNullOrWhiteSpace(model.ConsideracionesReserva) ? null : model.ConsideracionesReserva.Trim();
        model.GooglePlaceId = string.IsNullOrWhiteSpace(model.GooglePlaceId) ? null : model.GooglePlaceId.Trim();
        model.GoogleDepartamento = string.IsNullOrWhiteSpace(model.GoogleDepartamento) ? null : model.GoogleDepartamento.Trim();
        model.GoogleProvincia = string.IsNullOrWhiteSpace(model.GoogleProvincia) ? null : model.GoogleProvincia.Trim();
        model.GoogleDistrito = string.IsNullOrWhiteSpace(model.GoogleDistrito) ? null : model.GoogleDistrito.Trim();
        model.GoogleMapsUrl = string.IsNullOrWhiteSpace(model.GoogleMapsUrl) ? null : model.GoogleMapsUrl.Trim();
        model.FotoPrincipalUrl = string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) ? null : model.FotoPrincipalUrl.Trim();

        var fotos = (model.FotosUrlsCsv ?? string.Empty)
            .Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        model.FotosUrls = fotos;
        model.FotosUrlsCsv = fotos.Count == 0 ? null : string.Join(",", fotos);

        if (model.Latitud.HasValue && model.Longitud.HasValue && string.IsNullOrWhiteSpace(model.GoogleMapsUrl))
            model.GoogleMapsUrl = $"https://www.google.com/maps?q={model.Latitud.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)},{model.Longitud.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

        if (!string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && !Uri.IsWellFormedUriString(model.FotoPrincipalUrl, UriKind.Absolute))
            ModelState.AddModelError(nameof(model.FotoPrincipalUrl), "La foto principal debe ser una URL valida.");

        if (fotos.Any(url => !Uri.IsWellFormedUriString(url, UriKind.Absolute)))
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), "Todas las fotos de galeria deben ser URLs validas.");

        if (fotos.Count > MaxImagenesPorSede - 1)
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), $"La galeria permite maximo {MaxImagenesPorSede - 1} fotos alternativas.");

        var totalFotos = (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) ? 0 : 1) + fotos.Count;
        if (totalFotos > MaxImagenesPorSede)
            ModelState.AddModelError(nameof(model.FotosUrlsCsv), $"Solo se permiten {MaxImagenesPorSede} imagenes por sede.");

        if (string.IsNullOrWhiteSpace(model.FotoPrincipalUrl) && fotos.Count > 0)
            ModelState.AddModelError(nameof(model.FotoPrincipalUrl), "Debes tener una foto principal cuando registres fotos alternativas.");
    }

    private static void AplicarEliminacionImagenes(SedeFormViewModel model)
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

    private static List<string> ObtenerUrlsAEliminar(SedeFormViewModel model)
    {
        return (model.FotosEliminarUrls ?? new List<string>())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private async Task ProcesarCargaImagenesAsync(SedeFormViewModel model)
    {
        var archivos = (model.ImagenesArchivos ?? new List<IFormFile>())
            .Where(f => f is not null && f.Length > 0)
            .ToList();

        if (archivos.Count == 0)
            return;

        if (archivos.Count > MaxImagenesPorSede)
        {
            ModelState.AddModelError(nameof(model.ImagenesArchivos), $"Solo se permiten {MaxImagenesPorSede} imagenes por sede.");
            return;
        }

        try
        {
            var urls = await sedeImagenStorageService.UploadSedeImagenesAsync(
                model.NegocioId,
                model.Id > 0 ? model.Id : null,
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
            if (totalFinal > MaxImagenesPorSede)
            {
                ModelState.AddModelError(nameof(model.ImagenesArchivos), $"Solo se permiten {MaxImagenesPorSede} imagenes por sede. Ya tienes {urlsActuales.Count} registradas.");
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
}
