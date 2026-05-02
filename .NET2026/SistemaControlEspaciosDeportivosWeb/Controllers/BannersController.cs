using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class BannersController(
    ISportCenterStoredProcedureService spService,
    ISedeImagenStorageService sedeImagenStorageService) : Controller
{
    [Authorize(Roles = "OwnerPlataforma")]
    public async Task<IActionResult> Index(bool? soloActivos = null, int? editarId = null)
    {
        ViewData["PlatformShell"] = true;

        var vm = await BuildVmAsync(soloActivos, editarId);
        return View(vm);
    }

    [Authorize(Roles = "OwnerPlataforma")]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(BannerAdminFormViewModel form, bool? soloActivos = null)
    {
        ViewData["PlatformShell"] = true;

        form.TipoBanner = form.TipoBanner is >= 1 and <= 3 ? form.TipoBanner : (int)BannerTipo.Home;
        var esHome = form.TipoBanner == (int)BannerTipo.Home;
        var esLoginRegistro = form.TipoBanner is (int)BannerTipo.Login or (int)BannerTipo.Registro;
        var banners = await spService.BannersAdminListarAsync(null);
        if (!form.Id.HasValue && esLoginRegistro)
        {
            var existenteMismoTipo = banners
                .Where(x => x.TipoBanner == form.TipoBanner)
                .OrderBy(x => x.Orden)
                .ThenBy(x => x.Id)
                .FirstOrDefault();
            if (existenteMismoTipo is not null)
                form.Id = existenteMismoTipo.Id;
        }

        var bannerActual = form.Id.HasValue ? banners.FirstOrDefault(x => x.Id == form.Id.Value) : null;
        var imagenAnterior = string.IsNullOrWhiteSpace(bannerActual?.ImagenUrl) ? null : bannerActual!.ImagenUrl.Trim();
        var imagenMobileAnterior = string.IsNullOrWhiteSpace(bannerActual?.ImagenUrlMobile) ? null : bannerActual!.ImagenUrlMobile!.Trim();
        var imagenNueva = (string?)null;
        var imagenMobileNueva = (string?)null;

        if (form.ImagenArchivo is not null && form.ImagenArchivo.Length > 0)
        {
            try
            {
                if (esHome)
                {
                    imagenNueva = await sedeImagenStorageService.UploadBannerPublicoAsync(form.ImagenArchivo, HttpContext.RequestAborted);
                    form.ImagenUrl = imagenNueva;
                }
                else
                {
                    var orientacion = await ObtenerOrientacionAsync(form.ImagenArchivo, HttpContext.RequestAborted);
                    if (orientacion.EsHorizontal)
                    {
                        ModelState.AddModelError(nameof(form.ImagenArchivo), "En Login/Registro solo se permite imagen vertical.");
                    }
                    else
                    {
                        if (!orientacion.EsRelacionLoginValida)
                        {
                            ModelState.AddModelError(nameof(form.ImagenArchivo), "En Login/Registro la imagen debe tener proporcion 4:5 (ejemplo: 1080x1350).");
                        }
                        else
                        {
                            imagenMobileNueva = await sedeImagenStorageService.UploadBannerPublicoMobileAsync(form.ImagenArchivo, HttpContext.RequestAborted);
                            form.ImagenUrlMobile = imagenMobileNueva;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(nameof(form.ImagenArchivo), $"No se pudo subir la imagen: {ex.Message}");
            }
        }

        if (form.ImagenArchivoMobile is not null && form.ImagenArchivoMobile.Length > 0)
        {
            try
            {
                if (esHome)
                {
                    imagenMobileNueva = await sedeImagenStorageService.UploadBannerPublicoMobileAsync(form.ImagenArchivoMobile, HttpContext.RequestAborted);
                    form.ImagenUrlMobile = imagenMobileNueva;
                }
                else
                {
                    var orientacion = await ObtenerOrientacionAsync(form.ImagenArchivoMobile, HttpContext.RequestAborted);
                    if (orientacion.EsHorizontal)
                    {
                        ModelState.AddModelError(nameof(form.ImagenArchivoMobile), "En Login/Registro solo se permite imagen vertical.");
                    }
                    else
                    {
                        if (!orientacion.EsRelacionLoginValida)
                        {
                            ModelState.AddModelError(nameof(form.ImagenArchivoMobile), "En Login/Registro la imagen debe tener proporcion 4:5 (ejemplo: 1080x1350).");
                        }
                        else
                        {
                            imagenMobileNueva = await sedeImagenStorageService.UploadBannerPublicoMobileAsync(form.ImagenArchivoMobile, HttpContext.RequestAborted);
                            form.ImagenUrlMobile = imagenMobileNueva;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(nameof(form.ImagenArchivoMobile), $"No se pudo subir la imagen movil: {ex.Message}");
            }
        }

        form.Titulo = (form.Titulo ?? string.Empty).Trim();
        form.Subtitulo = string.IsNullOrWhiteSpace(form.Subtitulo) ? null : form.Subtitulo.Trim();
        form.Descripcion = string.IsNullOrWhiteSpace(form.Descripcion) ? null : form.Descripcion.Trim();
        form.BotonTexto = string.IsNullOrWhiteSpace(form.BotonTexto) ? null : form.BotonTexto.Trim();
        form.BotonUrl = string.IsNullOrWhiteSpace(form.BotonUrl) ? null : form.BotonUrl.Trim();
        form.ImagenUrl = string.IsNullOrWhiteSpace(form.ImagenUrl) ? imagenAnterior : form.ImagenUrl.Trim();
        form.ImagenUrlMobile = string.IsNullOrWhiteSpace(form.ImagenUrlMobile) ? imagenMobileAnterior : form.ImagenUrlMobile.Trim();

        if (esLoginRegistro && string.IsNullOrWhiteSpace(form.ImagenUrl) && !string.IsNullOrWhiteSpace(form.ImagenUrlMobile))
            form.ImagenUrl = form.ImagenUrlMobile;

        if (esHome && string.IsNullOrWhiteSpace(form.ImagenUrl))
            ModelState.AddModelError(nameof(form.ImagenArchivo), "Para Home debes cargar una imagen horizontal.");

        if (esLoginRegistro && string.IsNullOrWhiteSpace(form.ImagenUrlMobile))
            ModelState.AddModelError(nameof(form.ImagenArchivoMobile), "Para Login/Registro debes cargar una imagen vertical.");

        if (!ModelState.IsValid)
        {
            var urlsSubidas = new[] { imagenNueva, imagenMobileNueva }
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Cast<string>()
                .ToArray();
            if (urlsSubidas.Length > 0)
                await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsSubidas, HttpContext.RequestAborted);

            var primerError = ModelState.Values
                .SelectMany(v => v.Errors)
                .Select(e => e.ErrorMessage)
                .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));

            var vmError = await BuildVmAsync(soloActivos, null);
            vmError.Form = form;
            vmError.Error = string.IsNullOrWhiteSpace(primerError)
                ? "No se pudo guardar el banner. Revisa los campos."
                : $"No se pudo guardar el banner. {primerError}";
            return View(nameof(Index), vmError);
        }

        try
        {
            await spService.BannersAdminGuardarAsync(form, User.Identity?.Name ?? "sistema");

            var eliminarAnterior = !string.IsNullOrWhiteSpace(imagenAnterior) &&
                                   !string.IsNullOrWhiteSpace(imagenNueva) &&
                                   !string.Equals(imagenAnterior, imagenNueva, StringComparison.OrdinalIgnoreCase);
            var eliminarMobileAnterior = !string.IsNullOrWhiteSpace(imagenMobileAnterior) &&
                                         !string.IsNullOrWhiteSpace(imagenMobileNueva) &&
                                         !string.Equals(imagenMobileAnterior, imagenMobileNueva, StringComparison.OrdinalIgnoreCase);
            var urlsEliminar = new List<string>();
            if (eliminarAnterior) urlsEliminar.Add(imagenAnterior!);
            if (eliminarMobileAnterior) urlsEliminar.Add(imagenMobileAnterior!);
            if (urlsEliminar.Count > 0)
                await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsEliminar, HttpContext.RequestAborted);

            TempData["BannersOk"] = "Banner guardado correctamente.";
            return RedirectToAction(nameof(Index), new { soloActivos });
        }
        catch (Exception ex)
        {
            var urlsSubidas = new[] { imagenNueva, imagenMobileNueva }
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Cast<string>()
                .ToArray();
            if (urlsSubidas.Length > 0)
                await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsSubidas, HttpContext.RequestAborted);

            var vmError = await BuildVmAsync(soloActivos, null);
            vmError.Form = form;
            vmError.Error = ex.Message;
            return View(nameof(Index), vmError);
        }
    }

    [Authorize(Roles = "OwnerPlataforma")]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int id, string? imagenUrl, string? imagenUrlMobile, bool? soloActivos = null)
    {
        ViewData["PlatformShell"] = true;

        try
        {
            await spService.BannersAdminEliminarAsync(id, User.Identity?.Name ?? "sistema");
            var urlsEliminar = new[] { imagenUrl, imagenUrlMobile }
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (urlsEliminar.Length > 0)
                await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsEliminar, HttpContext.RequestAborted);

            TempData["BannersOk"] = "Banner eliminado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["BannersError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { soloActivos });
    }

    private async Task<BannersAdminIndexViewModel> BuildVmAsync(bool? soloActivos, int? editarId)
    {
        var banners = await spService.BannersAdminListarAsync(soloActivos);
        var editar = editarId.HasValue ? banners.FirstOrDefault(x => x.Id == editarId.Value) : null;
        var form = editar is null
            ? new BannerAdminFormViewModel()
            : new BannerAdminFormViewModel
            {
                Id = editar.Id,
                Titulo = editar.Titulo,
                Subtitulo = editar.Subtitulo,
                Descripcion = editar.Descripcion,
                BotonTexto = editar.BotonTexto,
                BotonUrl = editar.BotonUrl,
                ImagenUrl = editar.ImagenUrl,
                ImagenUrlMobile = editar.ImagenUrlMobile,
                TipoBanner = editar.TipoBanner,
                Orden = editar.Orden,
                Activo = editar.Activo,
                FechaInicio = editar.FechaInicio,
                FechaFin = editar.FechaFin
            };

        return new BannersAdminIndexViewModel
        {
            NegocioId = 0,
            NegocioNombre = "Plataforma",
            RolActual = "OwnerPlataforma",
            ModuloCodigo = "PLATFORM_BANNERS",
            ModuloNombre = "Banners Web",
            PuedeCrear = true,
            PuedeEditar = true,
            PuedeEliminar = true,
            SoloActivos = soloActivos,
            Banners = banners,
            Form = form,
            MensajeUi = TempData["BannersOk"]?.ToString(),
            Error = TempData["BannersError"]?.ToString()
        };
    }

    private static async Task<(bool EsHorizontal, bool EsRelacionLoginValida)> ObtenerOrientacionAsync(IFormFile archivo, CancellationToken cancellationToken)
    {
        await using var stream = archivo.OpenReadStream();
        using var image = await Image.LoadAsync<Rgba32>(stream, cancellationToken);
        image.Mutate(ctx => ctx.AutoOrient());
        if (image.Width <= 0 || image.Height <= 0)
            throw new InvalidOperationException($"No se pudo leer la imagen {archivo.FileName}.");

        var esHorizontal = image.Width >= image.Height;
        var ratio = image.Width / (decimal)image.Height;
        var ratioObjetivo = 1080m / 1350m; // 4:5
        var tolerancia = 0.03m; // +/-3%
        var esRelacionLoginValida = Math.Abs(ratio - ratioObjetivo) <= tolerancia;
        return (esHorizontal, esRelacionLoginValida);
    }
}
