using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class OrigenController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IOrigenRepository origenRepository) : Controller
{
    private const int TamanoPagina = 20;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var origenes = await origenRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, textoBusqueda, pagina, TamanoPagina, false, cancellationToken);
        var totalEmpresa = string.IsNullOrWhiteSpace(textoBusqueda)
            ? origenes.TotalRecords
            : (await origenRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, null, 1, 1, false, cancellationToken)).TotalRecords;
        var model = ConstruirViewModel(origenes.Items, null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalOrigenes = origenes.TotalRecords;
        model.PuedeCargarDefault = totalEmpresa == 0;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = origenes.TotalRecords
        };
        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idOrigen, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idOrigen, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CargarDefault(CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await origenRepository.CargarDefaultAsync(currentCompanyAccessor.EmpresaId.Value, User.Identity?.Name, cancellationToken);
        TempData["OrigenOk"] = "Origenes base cargados correctamente.";

        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(OrigenFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        if (!ModelState.IsValid)
        {
            var origenesError = await origenRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelError = ConstruirViewModel(origenesError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }

        try
        {
            await origenRepository.GuardarAsync(new GuardarOrigenRequest
            {
                IdOrigen = formulario.IdOrigen,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                CodigoOrigen = formulario.CodigoOrigen.Trim().ToUpperInvariant(),
                NombreOrigen = formulario.NombreOrigen.Trim(),
                ModuloOrigen = formulario.ModuloOrigen.Trim().ToUpperInvariant(),
                PermiteRegistroManual = formulario.PermiteRegistroManual,
                Estado = formulario.Estado,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["OrigenOk"] = formulario.IdOrigen.HasValue
                ? "Origen actualizado correctamente."
                : "Origen registrado correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var origenesError = await origenRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelError = ConstruirViewModel(origenesError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idOrigen, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var origenes = await origenRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
        var origenEditar = idOrigen.HasValue
            ? origenes.FirstOrDefault(x => x.IdOrigen == idOrigen.Value)
            : null;

        return View("Formulario", ConstruirViewModel(origenes, origenEditar));
    }

    private OrigenIndexViewModel ConstruirViewModel(IReadOnlyCollection<OrigenDto> origenes, OrigenDto? origenEditar)
    {
        var items = origenes
            .Select(x => new OrigenItemViewModel
            {
                IdOrigen = x.IdOrigen,
                CodigoOrigen = x.CodigoOrigen,
                NombreOrigen = x.NombreOrigen,
                ModuloOrigen = x.ModuloOrigen,
                PermiteRegistroManual = x.PermiteRegistroManual,
                Estado = x.Estado
            })
            .OrderBy(x => x.CodigoOrigen)
            .ToList();

        return new OrigenIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalOrigenes = items.Count,
            TotalActivos = items.Count(x => x.Estado),
            TotalManual = items.Count(x => x.PermiteRegistroManual),
            Origenes = items,
            Formulario = origenEditar is null
                ? new OrigenFormViewModel()
                : new OrigenFormViewModel
                {
                    IdOrigen = origenEditar.IdOrigen,
                    CodigoOrigen = origenEditar.CodigoOrigen,
                    NombreOrigen = origenEditar.NombreOrigen,
                    ModuloOrigen = origenEditar.ModuloOrigen,
                    PermiteRegistroManual = origenEditar.PermiteRegistroManual,
                    Estado = origenEditar.Estado
                }
        };
    }
}
