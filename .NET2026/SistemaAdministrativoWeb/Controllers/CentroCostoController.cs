using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CENTROCOSTO")]
public class CentroCostoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICentroCostoRepository centroCostoRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyuda = 100;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var resultado = await centroCostoRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            textoBusqueda,
            pagina,
            TamanoPagina,
            false,
            cancellationToken);

        var model = ConstruirViewModel(resultado.Items, null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalCentrosCosto = resultado.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = resultado.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idCentroCosto, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idCentroCosto, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("CENTROCOSTO", ModulePermissionOperation.Delete)]
    public async Task<IActionResult> Eliminar(int idCentroCosto, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await centroCostoRepository.EliminarAsync(currentCompanyAccessor.EmpresaId.Value, idCentroCosto, cancellationToken);
            TempData["CentroCostoOk"] = "Centro de costo eliminado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["CentroCostoError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { textoBusqueda, pagina });
    }

    [HttpGet]
    public async Task<IActionResult> BuscarAyuda(string? textoBusqueda = null, int tamanoPagina = TamanoAyuda, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var filtro = string.IsNullOrWhiteSpace(textoBusqueda) ? null : textoBusqueda.Trim();
        if (!string.IsNullOrWhiteSpace(filtro) && filtro.Length < 2)
        {
            filtro = null;
        }

        var resultado = await centroCostoRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            filtro,
            1,
            Math.Clamp(tamanoPagina, 1, TamanoAyuda),
            true,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idCentroCosto = x.IdCentroCosto,
                codigoCentroCosto = x.CodigoCentroCosto,
                nombreCentroCosto = x.NombreCentroCosto
            }),
            total = resultado.TotalRecords,
            limitado = resultado.TotalRecords > resultado.Items.Count
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModuleSavePermission("CENTROCOSTO", nameof(CentroCostoFormViewModel.IdCentroCosto))]
    public async Task<IActionResult> Guardar(CentroCostoFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        if (!ModelState.IsValid)
        {
            var centrosConError = await centroCostoRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(centrosConError, null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }

        try
        {
            await centroCostoRepository.GuardarAsync(new GuardarCentroCostoRequest
            {
                IdCentroCosto = formulario.IdCentroCosto,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                CodigoCentroCosto = formulario.CodigoCentroCosto.Trim().ToUpperInvariant(),
                NombreCentroCosto = formulario.NombreCentroCosto.Trim(),
                Estado = formulario.Estado
            }, cancellationToken);

            TempData["CentroCostoOk"] = formulario.IdCentroCosto.HasValue
                ? "Centro de costo actualizado correctamente."
                : "Centro de costo registrado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var centrosConError = await centroCostoRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(centrosConError, null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idCentroCosto, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var centros = await centroCostoRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
        var centroEditar = idCentroCosto.HasValue
            ? centros.FirstOrDefault(x => x.IdCentroCosto == idCentroCosto.Value)
            : null;

        return View("Formulario", ConstruirViewModel(centros, centroEditar));
    }

    private CentroCostoIndexViewModel ConstruirViewModel(IReadOnlyCollection<CentroCostoDto> centrosCosto, CentroCostoDto? centroEditar)
    {
        var items = centrosCosto
            .Select(x => new CentroCostoItemViewModel
            {
                IdCentroCosto = x.IdCentroCosto,
                CodigoCentroCosto = x.CodigoCentroCosto,
                NombreCentroCosto = x.NombreCentroCosto,
                Estado = x.Estado
            })
            .OrderBy(x => x.CodigoCentroCosto)
            .ToList();

        return new CentroCostoIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalCentrosCosto = items.Count,
            TotalActivos = items.Count(x => x.Estado),
            TotalInactivos = items.Count(x => !x.Estado),
            CentrosCosto = items,
            Formulario = centroEditar is null
                ? new CentroCostoFormViewModel()
                : new CentroCostoFormViewModel
                {
                    IdCentroCosto = centroEditar.IdCentroCosto,
                    CodigoCentroCosto = centroEditar.CodigoCentroCosto,
                    NombreCentroCosto = centroEditar.NombreCentroCosto,
                    Estado = centroEditar.Estado
                }
        };
    }
}
