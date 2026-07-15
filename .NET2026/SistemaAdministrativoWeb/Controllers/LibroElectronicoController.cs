using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize(Roles = "SuperAdmin,AdministradorEmpresa")]
[ModulePermission("LIBROELECTRONICO")]
public sealed class LibroElectronicoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IEmpresaRepository empresaRepository,
    ILibroElectronicoService libroElectronicoService,
    UserManager<IdentityUser> userManager,
    ILogger<LibroElectronicoController> logger) : Controller
{
    private const int TamanoPaginaPreview = 50;
    private const int TamanoPaginaHistorial = 10;

    [HttpGet]
    public async Task<IActionResult> Index(
        int? idEmpresa = null,
        short? anio = null,
        byte? mes = null,
        string? libroElectronico = null,
        string? moneda = null,
        string? operacion = null,
        int paginaPreview = 1,
        int paginaHistorial = 1,
        string? tokenDescarga = null,
        CancellationToken cancellationToken = default)
    {
        var model = await ConstruirModeloBaseAsync(idEmpresa, anio, mes, libroElectronico, moneda, paginaPreview, paginaHistorial, cancellationToken);
        model.TokenDescarga = tokenDescarga?.Trim() ?? string.Empty;
        model.PuedeDescargarTxt = !string.IsNullOrWhiteSpace(model.TokenDescarga) && model.PuedeDescargar;

        if (string.IsNullOrWhiteSpace(operacion))
        {
            return View(model);
        }

        try
        {
            var request = CrearRequest(model);
            var resultado = string.Equals(operacion, "validar", StringComparison.OrdinalIgnoreCase)
                || string.Equals(operacion, "observaciones", StringComparison.OrdinalIgnoreCase)
                ? await libroElectronicoService.ValidarAsync(request, model.EmpresaNombre, model.EmpresaRuc, paginaPreview, TamanoPaginaPreview, paginaHistorial, TamanoPaginaHistorial, cancellationToken)
                : await libroElectronicoService.ConsultarAsync(request, model.EmpresaNombre, model.EmpresaRuc, paginaPreview, TamanoPaginaPreview, paginaHistorial, TamanoPaginaHistorial, cancellationToken);

            AplicarResultado(model, resultado, operacion);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error consultando libros electrónicos para empresa {EmpresaId}, periodo {Anio}-{Mes}.", model.IdEmpresa, model.AnioSeleccionado, model.MesSeleccionado);
            model.MensajeError = "No se pudo procesar la consulta de libros electrónicos.";
        }

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GenerarTxt(LibroElectronicoViewModel model, CancellationToken cancellationToken = default)
    {
        var vista = await ConstruirModeloBaseAsync(model.IdEmpresa, model.AnioSeleccionado, model.MesSeleccionado, model.LibroElectronicoSeleccionado, model.MonedaSeleccionada, model.PaginaPreview, model.PaginaHistorial, cancellationToken);

        try
        {
            var request = CrearRequest(vista);
            var usuario = (await userManager.GetUserAsync(User))?.Email ?? User.Identity?.Name ?? "usuario";
            var resultado = await libroElectronicoService.GenerarAsync(request, vista.EmpresaNombre, vista.EmpresaRuc, usuario, vista.PaginaPreview, TamanoPaginaPreview, vista.PaginaHistorial, TamanoPaginaHistorial, cancellationToken);
            AplicarResultado(vista, resultado.Consulta, "generar");
            vista.TokenDescarga = resultado.TokenDescarga;
            vista.PuedeDescargarTxt = resultado.Generado && !string.IsNullOrWhiteSpace(resultado.TokenDescarga) && vista.PuedeDescargar;

            if (resultado.Generado)
            {
                vista.MensajeExito = resultado.Mensaje;
            }
            else
            {
                vista.MensajeError = resultado.Mensaje;
            }
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error generando libro electrónico para empresa {EmpresaId}, periodo {Anio}-{Mes}.", vista.IdEmpresa, vista.AnioSeleccionado, vista.MesSeleccionado);
            vista.MensajeError = "No se pudo generar el archivo TXT del libro electrónico.";
        }

        return View("Index", vista);
    }

    [HttpGet]
    public IActionResult DescargarTxt(string token)
    {
        if (!LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.DescargarTxt))
        {
            return Forbid();
        }

        var payload = libroElectronicoService.ObtenerDescarga(token, remover: true);
        if (payload is null)
        {
            TempData["LibroElectronicoError"] = "El archivo temporal ya no está disponible. Genérelo nuevamente.";
            return RedirectToAction(nameof(Index));
        }

        return File(payload.Content, "text/plain; charset=utf-8", payload.FileName);
    }

    private async Task<LibroElectronicoViewModel> ConstruirModeloBaseAsync(
        int? idEmpresa,
        short? anio,
        byte? mes,
        string? libroElectronico,
        string? moneda,
        int paginaPreview,
        int paginaHistorial,
        CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva && !User.IsInRole("SuperAdmin"))
        {
            throw new InvalidOperationException("No existe una empresa activa para la sesión actual.");
        }

        var userId = userManager.GetUserId(User);
        if (string.IsNullOrWhiteSpace(userId))
        {
            throw new InvalidOperationException("No se pudo identificar al usuario autenticado.");
        }

        var empresas = await empresaRepository.ListarPorUsuarioAsync(userId, cancellationToken);
        if (!currentCompanyAccessor.EmpresaId.HasValue)
        {
            throw new InvalidOperationException("No se encontró una empresa activa para la sesión actual.");
        }

        var empresaTrabajo = empresas.FirstOrDefault(x => x.IdEmpresa == currentCompanyAccessor.EmpresaId.Value);
        if (empresaTrabajo is null)
        {
            throw new InvalidOperationException("La empresa activa no está asignada al usuario actual.");
        }

        var anioTrabajo = anio ?? (short)DateTime.Today.Year;
        var mesTrabajo = mes is >= 1 and <= 12 ? mes.Value : (byte)DateTime.Today.Month;
        var libroTrabajo = PleLibroElectronicoCatalogo.Normalizar(libroElectronico);
        var monedaTrabajo = "PEN";

        var model = new LibroElectronicoViewModel
        {
            IdEmpresa = empresaTrabajo.IdEmpresa,
            EmpresaNombre = empresaTrabajo.RazonSocial,
            EmpresaRuc = empresaTrabajo.Ruc,
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            LibroElectronicoSeleccionado = libroTrabajo,
            MonedaSeleccionada = monedaTrabajo,
            PaginaPreview = Math.Max(1, paginaPreview),
            PaginaHistorial = Math.Max(1, paginaHistorial),
            TamanoPaginaPreview = TamanoPaginaPreview,
            TamanoPaginaHistorial = TamanoPaginaHistorial,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            MesesDisponibles = Enumerable.Range(1, 12)
                .Select(x => new MesOpcionViewModel
                {
                    Valor = (byte)x,
                    Nombre = new DateTime(2000, x, 1).ToString("MMMM")
                })
                .ToList(),
            LibrosDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = PleLibroElectronicoCatalogo.LibroDiario51, Texto = "5.1 - Libro Diario" },
                new OpcionCatalogoViewModel { Valor = PleLibroElectronicoCatalogo.LibroDiario52, Texto = "5.2 - Libro Diario Simplificado" },
                new OpcionCatalogoViewModel { Valor = PleLibroElectronicoCatalogo.LibroMayor61, Texto = "6.1 - Libro Mayor" }
            ],
            PuedeVer = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.Ver),
            PuedeConsultar = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.Consultar),
            PuedeValidar = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.Validar),
            PuedeGenerar = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.GenerarTxt),
            PuedeDescargar = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.DescargarTxt),
            PuedeVerHistorial = LibroElectronicoPermissions.TienePermiso(User, LibroElectronicoPermissions.VerHistorial)
        };

        if (TempData.TryGetValue("LibroElectronicoError", out var errorTemporal) && errorTemporal is string mensajeError)
        {
            model.MensajeError = mensajeError;
        }

        return model;
    }

    private static LibroElectronicoConsultaRequest CrearRequest(LibroElectronicoViewModel model)
    {
        return new LibroElectronicoConsultaRequest
        {
            IdEmpresa = model.IdEmpresa,
            Anio = model.AnioSeleccionado,
            Mes = model.MesSeleccionado,
            LibroElectronico = model.LibroElectronicoSeleccionado,
            Moneda = model.MonedaSeleccionada,
            Estado = "Todos",
            FechaDesde = null,
            FechaHasta = null
        };
    }

    private static void AplicarResultado(LibroElectronicoViewModel model, PleConsultaResultadoDto resultado, string operacion)
    {
        model.ConsultaEjecutada = true;
        model.ValidacionEjecutada = string.Equals(operacion, "validar", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operacion, "observaciones", StringComparison.OrdinalIgnoreCase)
            || string.Equals(operacion, "generar", StringComparison.OrdinalIgnoreCase);
        model.OperacionEjecutada = operacion;
        model.Resumen = resultado.Resumen;
        model.Validacion = resultado.Validacion;
        model.LibroDiario51Items = resultado.LibroDiario51Items;
        model.LibroDiario52Items = resultado.LibroDiario52Items;
        model.LibroMayor61Items = resultado.LibroMayor61Items;
        model.HistorialItems = resultado.Historial.Items;
        model.PreviewPaginacion = new PaginacionViewModel
        {
            PaginaActual = resultado.PaginaPreview,
            TamanoPagina = resultado.TamanoPaginaPreview,
            TotalRegistros = resultado.TotalRegistrosPreview
        };
        model.HistorialPaginacion = new PaginacionViewModel
        {
            PaginaActual = resultado.Historial.PageNumber,
            TamanoPagina = resultado.Historial.PageSize,
            TotalRegistros = resultado.Historial.TotalRecords
        };
    }
}
