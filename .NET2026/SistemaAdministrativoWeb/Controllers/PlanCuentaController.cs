using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class PlanCuentaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPlanCuentaRepository planCuentaRepository) : Controller
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

        var cuentas = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, textoBusqueda, pagina, TamanoPagina, false, cancellationToken);
        var model = ConstruirViewModel(cuentas.Items, null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalCuentas = cuentas.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = cuentas.TotalRecords
        };
        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idPlanCuenta, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idPlanCuenta, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(PlanCuentaFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        if (!ModelState.IsValid)
        {
            var cuentasConError = await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(cuentasConError, null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }

        try
        {
            await planCuentaRepository.GuardarAsync(new GuardarPlanCuentaRequest
            {
                IdPlanCuenta = formulario.IdPlanCuenta,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdPlanCuentaPadre = formulario.IdPlanCuentaPadre,
                CodigoCuenta = formulario.CodigoCuenta.Trim(),
                NombreCuenta = formulario.NombreCuenta.Trim(),
                NaturalezaSaldo = formulario.NaturalezaSaldo.Trim().ToUpperInvariant(),
                AceptaMovimiento = formulario.AceptaMovimiento,
                RequiereCentroCosto = formulario.RequiereCentroCosto,
                Estado = formulario.Estado,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["PlanCuentaOk"] = formulario.IdPlanCuenta.HasValue
                ? "Cuenta actualizada correctamente."
                : "Cuenta registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var cuentasConError = await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(cuentasConError, null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idPlanCuenta, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
        var cuentaEditar = idPlanCuenta.HasValue
            ? cuentas.FirstOrDefault(x => x.IdPlanCuenta == idPlanCuenta.Value)
            : null;

        var model = ConstruirViewModel(cuentas, cuentaEditar);
        return View("Formulario", model);
    }

    private PlanCuentaIndexViewModel ConstruirViewModel(IReadOnlyCollection<PlanCuentaDto> cuentas, PlanCuentaDto? cuentaEditar)
    {
        var items = cuentas
            .Select(x => new PlanCuentaItemViewModel
            {
                IdPlanCuenta = x.IdPlanCuenta,
                IdPlanCuentaPadre = x.IdPlanCuentaPadre,
                CodigoCuenta = x.CodigoCuenta,
                NombreCuenta = x.NombreCuenta,
                NivelCuenta = x.NivelCuenta,
                NaturalezaSaldo = x.NaturalezaSaldo,
                AceptaMovimiento = x.AceptaMovimiento,
                RequiereCentroCosto = x.RequiereCentroCosto,
                Estado = x.Estado
            })
            .OrderBy(x => x.CodigoCuenta)
            .ToList();

        return new PlanCuentaIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalCuentas = items.Count,
            TotalMovimiento = items.Count(x => x.AceptaMovimiento),
            TotalActivas = items.Count(x => x.Estado),
            Cuentas = items,
            CuentasPadre = cuentaEditar is null
                ? items
                : items.Where(x => x.IdPlanCuenta != cuentaEditar.IdPlanCuenta).ToList(),
            Formulario = cuentaEditar is null
                ? new PlanCuentaFormViewModel()
                : new PlanCuentaFormViewModel
                {
                    IdPlanCuenta = cuentaEditar.IdPlanCuenta,
                    IdPlanCuentaPadre = cuentaEditar.IdPlanCuentaPadre,
                    CodigoCuenta = cuentaEditar.CodigoCuenta,
                    NombreCuenta = cuentaEditar.NombreCuenta,
                    NaturalezaSaldo = cuentaEditar.NaturalezaSaldo,
                    AceptaMovimiento = cuentaEditar.AceptaMovimiento,
                    RequiereCentroCosto = cuentaEditar.RequiereCentroCosto,
                    Estado = cuentaEditar.Estado
                }
        };
    }
}
