using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize(Roles = "OwnerPlataforma")]
public class AnunciosController(
    ISportCenterStoredProcedureService spService,
    ISedeImagenStorageService sedeImagenStorageService) : Controller
{
    private static readonly (string Key, string Label, string? Desc, Func<PopupPromocionConfigViewModel, string> GetValue)[] PopupConfigMap =
    [
        ("POPUP_PROMO_AUTO_ENABLED", "Popup promociones automatico", "Controla si el popup de promociones se abre automaticamente en el home publico.", x => x.ActivarPopupAutomatico ? "1" : "0"),
        ("POPUP_PROMO_DELAY_SECONDS", "Popup promociones espera", "Segundos de espera antes de mostrar el popup de promociones en el home publico.", x => x.SegundosEsperaAntesDeMostrar.ToString()),
        ("POPUP_PROMO_AUTOPLAY_ENABLED", "Popup promociones autoplay", "Controla si el slider del popup rota automaticamente.", x => x.ActivarAutoplaySlider ? "1" : "0"),
        ("POPUP_PROMO_AUTOPLAY_MS", "Popup promociones velocidad", "Velocidad del autoplay del slider de promociones en milisegundos.", x => x.VelocidadAutoplayMs.ToString()),
        ("POPUP_PROMO_SHOW_ARROWS", "Popup promociones flechas", "Controla si el slider del popup muestra flechas laterales.", x => x.MostrarFlechas ? "1" : "0"),
        ("POPUP_PROMO_SHOW_INDICATORS", "Popup promociones indicadores", "Controla si el slider del popup muestra indicadores inferiores.", x => x.MostrarIndicadores ? "1" : "0")
    ];

    public async Task<IActionResult> Index(bool? soloActivos = null, int? editarId = null)
    {
        ViewData["PlatformShell"] = true;
        var vm = await BuildVmAsync(soloActivos, editarId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(PopupPromocionAdminFormViewModel form, bool? soloActivos = null)
    {
        ViewData["PlatformShell"] = true;

        var anuncios = await spService.PopupPromocionesAdminListarAsync(null);
        var anuncioActual = form.IdPopupPromocion.HasValue
            ? anuncios.FirstOrDefault(x => x.IdPopupPromocion == form.IdPopupPromocion.Value)
            : null;
        var imagenAnterior = string.IsNullOrWhiteSpace(anuncioActual?.ImagenUrl) ? null : anuncioActual!.ImagenUrl.Trim();
        var imagenNueva = (string?)null;

        if (form.ImagenArchivo is not null && form.ImagenArchivo.Length > 0)
        {
            try
            {
                form.Orientacion = NormalizarOrientacion(form.Orientacion);
                imagenNueva = await sedeImagenStorageService.UploadBannerAnuncioAsync(
                    form.ImagenArchivo,
                    esHorizontal: string.Equals(form.Orientacion, PopupPromocionPublicoViewModel.OrientacionHorizontal, StringComparison.OrdinalIgnoreCase),
                    HttpContext.RequestAborted);
                form.ImagenUrl = imagenNueva;
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(nameof(form.ImagenArchivo), $"No se pudo subir la imagen: {ex.Message}");
            }
        }

        form.Titulo = (form.Titulo ?? string.Empty).Trim();
        form.Subtitulo = string.IsNullOrWhiteSpace(form.Subtitulo) ? null : form.Subtitulo.Trim();
        form.Descripcion = string.IsNullOrWhiteSpace(form.Descripcion) ? null : form.Descripcion.Trim();
        form.Orientacion = NormalizarOrientacion(form.Orientacion);
        form.ImagenUrl = string.IsNullOrWhiteSpace(form.ImagenUrl) ? imagenAnterior : form.ImagenUrl.Trim();
        form.TextoBoton = string.IsNullOrWhiteSpace(form.TextoBoton) ? null : form.TextoBoton.Trim();
        form.UrlBoton = string.IsNullOrWhiteSpace(form.UrlBoton) ? null : form.UrlBoton.Trim();
        form.UrlImagen = string.IsNullOrWhiteSpace(form.UrlImagen) ? null : form.UrlImagen.Trim();

        if (anuncioActual is not null &&
            !string.Equals(anuncioActual.Orientacion, form.Orientacion, StringComparison.OrdinalIgnoreCase) &&
            (form.ImagenArchivo is null || form.ImagenArchivo.Length <= 0))
        {
            ModelState.AddModelError(nameof(form.ImagenArchivo), "Si cambias la orientacion del anuncio, debes cargar una nueva imagen del tipo seleccionado.");
        }

        ValidarFormulario(form);

        if (!ModelState.IsValid)
        {
            if (!string.IsNullOrWhiteSpace(imagenNueva))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenNueva], HttpContext.RequestAborted);

            var vmError = await BuildVmAsync(soloActivos, null);
            vmError.Form = form;
            vmError.Error = "No se pudo guardar el anuncio. Revisa los campos.";
            return View(nameof(Index), vmError);
        }

        try
        {
            await spService.PopupPromocionesAdminGuardarAsync(form, User.Identity?.Name ?? "owner-platform");

            var debeEliminarImagenAnterior = !string.IsNullOrWhiteSpace(imagenAnterior) &&
                                             !string.IsNullOrWhiteSpace(imagenNueva) &&
                                             !string.Equals(imagenAnterior, imagenNueva, StringComparison.OrdinalIgnoreCase);
            if (debeEliminarImagenAnterior)
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenAnterior!], HttpContext.RequestAborted);

            TempData["PopupPromocionesOk"] = "Anuncio guardado correctamente.";
            return RedirectToAction(nameof(Index), new { soloActivos });
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrWhiteSpace(imagenNueva))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenNueva], HttpContext.RequestAborted);

            var vmError = await BuildVmAsync(soloActivos, null);
            vmError.Form = form;
            vmError.Error = ex.Message;
            return View(nameof(Index), vmError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idPopupPromocion, string? imagenUrl, bool? soloActivos = null)
    {
        ViewData["PlatformShell"] = true;

        try
        {
            await spService.PopupPromocionesAdminEliminarAsync(idPopupPromocion, User.Identity?.Name ?? "owner-platform");
            if (!string.IsNullOrWhiteSpace(imagenUrl))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenUrl.Trim()], HttpContext.RequestAborted);

            TempData["PopupPromocionesOk"] = "Anuncio eliminado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PopupPromocionesError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { soloActivos });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CambiarEstado(int idPopupPromocion, bool activo, bool? soloActivos = null)
    {
        ViewData["PlatformShell"] = true;

        try
        {
            await spService.PopupPromocionesAdminCambiarEstadoAsync(idPopupPromocion, activo, User.Identity?.Name ?? "owner-platform");
            TempData["PopupPromocionesOk"] = activo ? "Anuncio activado correctamente." : "Anuncio desactivado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PopupPromocionesError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { soloActivos });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarConfiguracion(PopupPromocionConfigViewModel config, bool? soloActivos = null, int? editarId = null)
    {
        ViewData["PlatformShell"] = true;

        if (!ModelState.IsValid)
        {
            var vmError = await BuildVmAsync(soloActivos, editarId);
            vmError.Config = config;
            vmError.Error = "No se pudo guardar la configuracion del popup. Revisa los campos.";
            return View(nameof(Index), vmError);
        }

        try
        {
            var usuario = User.Identity?.Name ?? "owner-platform";
            foreach (var item in PopupConfigMap)
            {
                await spService.ParametrosGlobalesUpsertValorAsync(item.Key, item.Desc, item.GetValue(config), usuario);
            }

            TempData["PopupPromocionesOk"] = "Configuracion del popup actualizada.";
            return RedirectToAction(nameof(Index), new { soloActivos });
        }
        catch (Exception ex)
        {
            var vmError = await BuildVmAsync(soloActivos, editarId);
            vmError.Config = config;
            vmError.Error = $"No se pudo guardar la configuracion: {ex.Message}";
            return View(nameof(Index), vmError);
        }
    }

    private async Task<PopupPromocionesAdminIndexViewModel> BuildVmAsync(bool? soloActivos, int? editarId)
    {
        var anuncios = await spService.PopupPromocionesAdminListarAsync(soloActivos);
        var editar = editarId.HasValue
            ? anuncios.FirstOrDefault(x => x.IdPopupPromocion == editarId.Value)
            : null;
        var form = editar is null
            ? new PopupPromocionAdminFormViewModel()
            : new PopupPromocionAdminFormViewModel
            {
                IdPopupPromocion = editar.IdPopupPromocion,
                Titulo = editar.Titulo,
                Subtitulo = editar.Subtitulo,
                Descripcion = editar.Descripcion,
                ImagenUrl = editar.ImagenUrl,
                Orientacion = editar.Orientacion,
                TextoBoton = editar.TextoBoton,
                UrlBoton = editar.UrlBoton,
                UrlImagen = editar.UrlImagen,
                Orden = editar.Orden,
                Activo = editar.Activo,
                FechaInicio = editar.FechaInicio,
                FechaFin = editar.FechaFin,
                AbrirNuevaPestana = editar.AbrirNuevaPestana
            };

        return new PopupPromocionesAdminIndexViewModel
        {
            NegocioId = 0,
            NegocioNombre = "Plataforma",
            RolActual = "OwnerPlataforma",
            ModuloCodigo = "PLATFORM_POPUP_PROMOS",
            ModuloNombre = "Anuncios",
            PuedeCrear = true,
            PuedeEditar = true,
            PuedeEliminar = true,
            Anuncios = anuncios,
            Form = form,
            Config = await CargarConfiguracionAsync(),
            MensajeUi = TempData["PopupPromocionesOk"]?.ToString(),
            Error = TempData["PopupPromocionesError"]?.ToString()
        };
    }

    private async Task<PopupPromocionConfigViewModel> CargarConfiguracionAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);
        return new PopupPromocionConfigViewModel
        {
            ActivarPopupAutomatico = LeerBool(await Get("POPUP_PROMO_AUTO_ENABLED"), true),
            SegundosEsperaAntesDeMostrar = LeerEntero(await Get("POPUP_PROMO_DELAY_SECONDS"), 1, 0, 30),
            ActivarAutoplaySlider = LeerBool(await Get("POPUP_PROMO_AUTOPLAY_ENABLED"), true),
            VelocidadAutoplayMs = LeerEntero(await Get("POPUP_PROMO_AUTOPLAY_MS"), 4500, 1000, 20000),
            MostrarFlechas = LeerBool(await Get("POPUP_PROMO_SHOW_ARROWS"), true),
            MostrarIndicadores = LeerBool(await Get("POPUP_PROMO_SHOW_INDICATORS"), true)
        };
    }

    private void ValidarFormulario(PopupPromocionAdminFormViewModel form)
    {
        if (!string.Equals(form.Orientacion, PopupPromocionPublicoViewModel.OrientacionVertical, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(form.Orientacion, PopupPromocionPublicoViewModel.OrientacionHorizontal, StringComparison.OrdinalIgnoreCase))
            ModelState.AddModelError(nameof(form.Orientacion), "Selecciona una orientacion valida para el anuncio.");

        if (string.IsNullOrWhiteSpace(form.ImagenUrl))
            ModelState.AddModelError(nameof(form.ImagenArchivo), "Debes cargar una imagen para el anuncio.");

        if (form.FechaInicio.HasValue && form.FechaFin.HasValue && form.FechaFin.Value < form.FechaInicio.Value)
            ModelState.AddModelError(nameof(form.FechaFin), "La fecha fin no puede ser menor a la fecha inicio.");

        if (!string.IsNullOrWhiteSpace(form.TextoBoton) && string.IsNullOrWhiteSpace(form.UrlBoton))
            ModelState.AddModelError(nameof(form.UrlBoton), "Ingresa la URL del boton.");

        if (string.IsNullOrWhiteSpace(form.TextoBoton) && !string.IsNullOrWhiteSpace(form.UrlBoton))
            ModelState.AddModelError(nameof(form.TextoBoton), "Ingresa el texto del boton.");

        if (!EsUrlValida(form.UrlBoton))
            ModelState.AddModelError(nameof(form.UrlBoton), "Ingresa una URL valida para el boton.");

        if (!EsUrlValida(form.UrlImagen))
            ModelState.AddModelError(nameof(form.UrlImagen), "Ingresa una URL valida para la imagen.");
    }

    private static bool EsUrlValida(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return true;

        var url = valor.Trim();
        if (url.StartsWith("/", StringComparison.Ordinal) || url.StartsWith("#", StringComparison.Ordinal))
            return true;

        return Uri.TryCreate(url, UriKind.Absolute, out var uri) &&
               (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
    }

    private static bool LeerBool(string? valor, bool fallback)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return fallback;

        valor = valor.Trim();
        if (valor == "1")
            return true;
        if (valor == "0")
            return false;

        return bool.TryParse(valor, out var parsed) ? parsed : fallback;
    }

    private static int LeerEntero(string? valor, int fallback, int min, int max)
    {
        if (!int.TryParse((valor ?? string.Empty).Trim(), out var parsed))
            return fallback;

        return Math.Clamp(parsed, min, max);
    }

    private static string NormalizarOrientacion(string? valor)
    {
        return string.Equals((valor ?? string.Empty).Trim(), PopupPromocionPublicoViewModel.OrientacionHorizontal, StringComparison.OrdinalIgnoreCase)
            ? PopupPromocionPublicoViewModel.OrientacionHorizontal
            : PopupPromocionPublicoViewModel.OrientacionVertical;
    }
}
