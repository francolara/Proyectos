using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

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

        var banners = await spService.BannersAdminListarAsync(null);
        var bannerActual = form.Id.HasValue ? banners.FirstOrDefault(x => x.Id == form.Id.Value) : null;
        var imagenAnterior = string.IsNullOrWhiteSpace(bannerActual?.ImagenUrl) ? null : bannerActual!.ImagenUrl.Trim();
        var imagenMobileAnterior = string.IsNullOrWhiteSpace(bannerActual?.ImagenUrlMobile) ? null : bannerActual!.ImagenUrlMobile!.Trim();
        var imagenNueva = (string?)null;
        var imagenMobileNueva = (string?)null;

        if (form.ImagenArchivo is not null && form.ImagenArchivo.Length > 0)
        {
            try
            {
                imagenNueva = await sedeImagenStorageService.UploadBannerPublicoAsync(form.ImagenArchivo, HttpContext.RequestAborted);
                form.ImagenUrl = imagenNueva;
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
                imagenMobileNueva = await sedeImagenStorageService.UploadBannerPublicoMobileAsync(form.ImagenArchivoMobile, HttpContext.RequestAborted);
                form.ImagenUrlMobile = imagenMobileNueva;
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
        form.TipoBanner = form.TipoBanner is >= 1 and <= 3 ? form.TipoBanner : (int)BannerTipo.Home;

        if (string.IsNullOrWhiteSpace(form.ImagenUrl))
            ModelState.AddModelError(nameof(form.ImagenArchivo), "Debes cargar una imagen para el banner.");

        if (!ModelState.IsValid)
        {
            var urlsSubidas = new[] { imagenNueva, imagenMobileNueva }
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Cast<string>()
                .ToArray();
            if (urlsSubidas.Length > 0)
                await sedeImagenStorageService.DeleteSedeImagenesAsync(urlsSubidas, HttpContext.RequestAborted);

            var vmError = await BuildVmAsync(soloActivos, null);
            vmError.Form = form;
            vmError.Error = "No se pudo guardar el banner. Revisa los campos.";
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
}
