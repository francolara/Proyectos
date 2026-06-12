using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Authorize(Roles = "OwnerPlataforma")]
public class BoletinesAdminController(
    ISportCenterStoredProcedureService spService,
    ISedeImagenStorageService sedeImagenStorageService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(
        bool? soloActivos = null,
        string? tipoRegistro = null,
        string? codigoDepartamento = null,
        string? codigoProvincia = null,
        string? codigoUbigeo = null,
        string? zona = null,
        int? anio = null,
        int? mes = null,
        int pagina = 1,
        int? editarId = null)
    {
        ViewData["PlatformShell"] = true;
        var vm = await BuildVmAsync(
            soloActivos,
            tipoRegistro,
            codigoDepartamento,
            codigoProvincia,
            codigoUbigeo,
            zona,
            anio,
            mes,
            pagina,
            editarId);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(BoletinesDeportivosAdminIndexViewModel vm)
    {
        ViewData["PlatformShell"] = true;

        var usuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "owner-platform";
        var usuario = User.Identity?.Name ?? "owner-platform";

        vm.TipoRegistro = NormalizarTipoRegistro(vm.TipoRegistro);
        vm.CodigoDepartamento = NormalizarCodigo(vm.CodigoDepartamento);
        vm.CodigoProvincia = NormalizarCodigo(vm.CodigoProvincia);
        vm.CodigoUbigeo = NormalizarCodigo(vm.CodigoUbigeo);
        vm.Zona = NormalizarTexto(vm.Zona);
        vm.Form.UsuarioId = usuarioId;
        vm.Form.TipoRegistro = "A";
        vm.Form.EsAdministradorCarga = true;
        vm.Form.Titulo = NormalizarTexto(vm.Form.Titulo);
        vm.Form.Descripcion = NormalizarTexto(vm.Form.Descripcion);
        vm.Form.CodigoUbigeo = NormalizarCodigo(vm.Form.CodigoUbigeo) ?? string.Empty;
        vm.Form.CodigoDepartamento = NormalizarCodigo(vm.Form.CodigoDepartamento);
        vm.Form.CodigoProvincia = NormalizarCodigo(vm.Form.CodigoProvincia);
        vm.Form.ImagenUrl = NormalizarTexto(vm.Form.ImagenUrl) ?? string.Empty;

        string? imagenNueva = null;
        if (vm.Form.ImagenArchivo is not null && vm.Form.ImagenArchivo.Length > 0)
        {
            try
            {
                imagenNueva = await sedeImagenStorageService.UploadBoletinDeportivoAsync(vm.Form.ImagenArchivo, HttpContext.RequestAborted);
                vm.Form.ImagenUrl = imagenNueva ?? string.Empty;
                ModelState.Remove("Form.ImagenUrl");
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("Form.ImagenArchivo", $"No se pudo subir el boletin: {ex.Message}");
            }
        }

        if (string.IsNullOrWhiteSpace(vm.Form.ImagenUrl))
            ModelState.AddModelError("Form.ImagenArchivo", "Debes cargar la imagen del boletin.");

        if (!ModelState.IsValid)
        {
            if (!string.IsNullOrWhiteSpace(imagenNueva))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenNueva], HttpContext.RequestAborted);

            await CargarCombosFiltrosAsync(vm);
            await CargarCombosFormAsync(vm.Form);
            var listado = await spService.BoletinesDeportivosAdminListarAsync(
                vm.SoloActivos,
                vm.TipoRegistro,
                vm.CodigoDepartamento,
                vm.CodigoProvincia,
                vm.CodigoUbigeo,
                vm.Zona,
                vm.Anio,
                vm.Mes,
                vm.Pagina,
                vm.TamanoPagina);
            vm.Boletines = listado.Boletines;
            vm.TotalRegistros = listado.TotalRegistros;
            vm.TotalPaginas = vm.TotalRegistros > 0
                ? (int)Math.Ceiling(vm.TotalRegistros / (double)vm.TamanoPagina)
                : 1;
            vm.Error = "No se pudo guardar el boletin. Revisa los campos.";
            CompletarContexto(vm);
            return View(nameof(Index), vm);
        }

        try
        {
            await spService.BoletinesDeportivosGuardarAsync(vm.Form, usuario);
            TempData["BoletinesAdminOk"] = vm.Form.IdBoletin.HasValue
                ? "Boletin actualizado correctamente."
                : "Boletin publicado correctamente.";
            return RedirectToAction(nameof(Index), new
            {
                soloActivos = vm.SoloActivos,
                tipoRegistro = vm.TipoRegistro,
                codigoDepartamento = vm.CodigoDepartamento,
                codigoProvincia = vm.CodigoProvincia,
                codigoUbigeo = vm.CodigoUbigeo,
                zona = vm.Zona,
                anio = vm.Anio,
                mes = vm.Mes,
                pagina = vm.Pagina
            });
        }
        catch (Exception ex)
        {
            if (!string.IsNullOrWhiteSpace(imagenNueva))
                await sedeImagenStorageService.DeleteSedeImagenesAsync([imagenNueva], HttpContext.RequestAborted);

            await CargarCombosFiltrosAsync(vm);
            await CargarCombosFormAsync(vm.Form);
            var listado = await spService.BoletinesDeportivosAdminListarAsync(
                vm.SoloActivos,
                vm.TipoRegistro,
                vm.CodigoDepartamento,
                vm.CodigoProvincia,
                vm.CodigoUbigeo,
                vm.Zona,
                vm.Anio,
                vm.Mes,
                vm.Pagina,
                vm.TamanoPagina);
            vm.Boletines = listado.Boletines;
            vm.TotalRegistros = listado.TotalRegistros;
            vm.TotalPaginas = vm.TotalRegistros > 0
                ? (int)Math.Ceiling(vm.TotalRegistros / (double)vm.TamanoPagina)
                : 1;
            vm.Error = ex.Message;
            CompletarContexto(vm);
            return View(nameof(Index), vm);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CambiarEstado(
        int idBoletin,
        int activo,
        bool? soloActivos = null,
        string? tipoRegistro = null,
        string? codigoDepartamento = null,
        string? codigoProvincia = null,
        string? codigoUbigeo = null,
        string? zona = null,
        int? anio = null,
        int? mes = null,
        int pagina = 1)
    {
        ViewData["PlatformShell"] = true;
        var activar = activo == 1;

        try
        {
            var estadoActual = await spService.BoletinesDeportivosCambiarEstadoAsync(idBoletin, activar, User.Identity?.Name ?? "owner-platform");
            TempData["BoletinesAdminOk"] = estadoActual
                ? "Boletin activado correctamente."
                : "Boletin desactivado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["BoletinesAdminError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new
        {
            soloActivos,
            tipoRegistro = NormalizarTipoRegistro(tipoRegistro),
            codigoDepartamento = NormalizarCodigo(codigoDepartamento),
            codigoProvincia = NormalizarCodigo(codigoProvincia),
            codigoUbigeo = NormalizarCodigo(codigoUbigeo),
            zona = NormalizarTexto(zona),
            anio,
            mes,
            pagina
        });
    }

    private async Task<BoletinesDeportivosAdminIndexViewModel> BuildVmAsync(
        bool? soloActivos,
        string? tipoRegistro,
        string? codigoDepartamento,
        string? codigoProvincia,
        string? codigoUbigeo,
        string? zona,
        int? anio,
        int? mes,
        int pagina,
        int? editarId)
    {
        var resumen = await spService.BoletinesDeportivosAdminResumenAsync();
        var vm = new BoletinesDeportivosAdminIndexViewModel
        {
            TotalBoletines = resumen.TotalBoletines,
            TotalActivos = resumen.TotalActivos,
            TotalInactivos = resumen.TotalInactivos,
            TotalUsuarios = resumen.TotalUsuarios,
            TotalPlataforma = resumen.TotalPlataforma,
            Pagina = pagina < 1 ? 1 : pagina,
            SoloActivos = soloActivos,
            TipoRegistro = NormalizarTipoRegistro(tipoRegistro),
            CodigoDepartamento = NormalizarCodigo(codigoDepartamento),
            CodigoProvincia = NormalizarCodigo(codigoProvincia),
            CodigoUbigeo = NormalizarCodigo(codigoUbigeo),
            Zona = NormalizarTexto(zona),
            Anio = anio,
            Mes = mes,
            MensajeUi = TempData["BoletinesAdminOk"]?.ToString(),
            Error = TempData["BoletinesAdminError"]?.ToString()
        };

        var listado = await spService.BoletinesDeportivosAdminListarAsync(
            vm.SoloActivos,
            vm.TipoRegistro,
            vm.CodigoDepartamento,
            vm.CodigoProvincia,
            vm.CodigoUbigeo,
            vm.Zona,
            vm.Anio,
            vm.Mes,
            vm.Pagina,
            vm.TamanoPagina);
        vm.Boletines = listado.Boletines;
        vm.TotalRegistros = listado.TotalRegistros;
        vm.TotalPaginas = vm.TotalRegistros > 0
            ? (int)Math.Ceiling(vm.TotalRegistros / (double)vm.TamanoPagina)
            : 1;
        if (vm.Pagina > vm.TotalPaginas)
        {
            vm.Pagina = vm.TotalPaginas;
            listado = await spService.BoletinesDeportivosAdminListarAsync(
                vm.SoloActivos,
                vm.TipoRegistro,
                vm.CodigoDepartamento,
                vm.CodigoProvincia,
                vm.CodigoUbigeo,
                vm.Zona,
                vm.Anio,
                vm.Mes,
                vm.Pagina,
                vm.TamanoPagina);
            vm.Boletines = listado.Boletines;
        }

        var form = new BoletinDeportivoGuardarViewModel
        {
            UsuarioId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "owner-platform",
            TipoRegistro = "A",
            Activo = true,
            EsAdministradorCarga = true,
            FechaEvento = DateOnly.FromDateTime(DateTime.Today)
        };

        if (editarId.HasValue)
        {
            var item = vm.Boletines.FirstOrDefault(x => x.IdBoletin == editarId.Value);
            if (item is not null)
            {
                var ubigeo = await spService.UbigeoObtenerPorCodigoAsync(item.CodigoUbigeo);
                form = new BoletinDeportivoGuardarViewModel
                {
                    IdBoletin = item.IdBoletin,
                    UsuarioId = item.UsuarioId,
                    Titulo = item.Titulo,
                    Descripcion = item.Descripcion,
                    FechaEvento = item.FechaEvento,
                    CodigoUbigeo = item.CodigoUbigeo,
                    ImagenUrl = item.ImagenUrl,
                    Activo = item.Activo,
                    TipoRegistro = item.TipoRegistro,
                    EsAdministradorCarga = true,
                    CodigoDepartamento = ubigeo?.CodigoDepartamento,
                    CodigoProvincia = ubigeo?.CodigoProvincia,
                    Zona = item.Zona
                };
            }
        }

        vm.Form = form;
        await CargarCombosFiltrosAsync(vm);
        await CargarCombosFormAsync(vm.Form);
        CompletarContexto(vm);
        return vm;
    }

    private async Task CargarCombosFiltrosAsync(BoletinesDeportivosAdminIndexViewModel vm)
    {
        vm.Departamentos = await spService.UbigeoDepartamentosListarAsync();
        vm.Provincias = !string.IsNullOrWhiteSpace(vm.CodigoDepartamento) && vm.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(vm.CodigoDepartamento)
            : new List<SelectListItem>();
        vm.Distritos = !string.IsNullOrWhiteSpace(vm.CodigoProvincia) && vm.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(vm.CodigoProvincia)
            : new List<SelectListItem>();
        vm.Zonas = await spService.UbigeoZonasListarAsync(vm.CodigoDepartamento, vm.CodigoProvincia);

        var anioActual = DateTime.Today.Year;
        vm.Anios =
        [
            .. Enumerable.Range(anioActual - 1, 5)
                .Select(x => new SelectListItem(x.ToString(), x.ToString(), vm.Anio == x))
        ];
        vm.Meses =
        [
            .. Enumerable.Range(1, 12)
                .Select(x => new SelectListItem(
                    new DateTime(2000, x, 1).ToString("MMMM"),
                    x.ToString(),
                    vm.Mes == x))
        ];
    }

    private async Task CargarCombosFormAsync(BoletinDeportivoGuardarViewModel form)
    {
        form.Departamentos = await spService.UbigeoDepartamentosListarAsync();
        form.Provincias = !string.IsNullOrWhiteSpace(form.CodigoDepartamento) && form.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(form.CodigoDepartamento)
            : new List<SelectListItem>();
        form.Distritos = !string.IsNullOrWhiteSpace(form.CodigoProvincia) && form.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(form.CodigoProvincia)
            : new List<SelectListItem>();
        form.Zonas = await spService.UbigeoZonasListarAsync(form.CodigoDepartamento, form.CodigoProvincia);
    }

    private static void CompletarContexto(BoletinesDeportivosAdminIndexViewModel vm)
    {
        vm.NegocioId = 0;
        vm.NegocioNombre = "Plataforma";
        vm.RolActual = "OwnerPlataforma";
        vm.ModuloCodigo = "PLATFORM_BOLETINES";
        vm.ModuloNombre = "Boletines";
        vm.PuedeCrear = true;
        vm.PuedeEditar = true;
        vm.PuedeEliminar = false;
    }

    private static string? NormalizarTexto(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();

    private static string? NormalizarCodigo(string? valor)
        => string.IsNullOrWhiteSpace(valor) ? null : valor.Trim();

    private static string? NormalizarTipoRegistro(string? valor)
    {
        var normalizado = string.IsNullOrWhiteSpace(valor) ? null : valor.Trim().ToUpperInvariant();
        return normalizado is "U" or "A" ? normalizado : null;
    }
}
