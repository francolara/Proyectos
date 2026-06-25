using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class PlanCuentaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPlanCuentaRepository planCuentaRepository,
    IMonedaRepository monedaRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, byte? nivelCuenta = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var nivelCuentaTrabajo = nivelCuenta is >= 1 and <= 5 ? nivelCuenta : null;
        var cuentas = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, textoBusqueda, nivelCuentaTrabajo, pagina, TamanoPagina, false, false, cancellationToken);
        var totalEmpresa = string.IsNullOrWhiteSpace(textoBusqueda) && !nivelCuentaTrabajo.HasValue
            ? cuentas.TotalRecords
            : (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, null, null, 1, 1, false, false, cancellationToken)).TotalRecords;
        var model = ConstruirViewModel(cuentas.Items, null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.NivelCuentaFiltro = nivelCuentaTrabajo;
        model.TotalCuentas = cuentas.TotalRecords;
        model.PuedeCargarDefault = totalEmpresa == 0;
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

    [HttpGet]
    public async Task<IActionResult> BuscarAyuda(string? textoBusqueda = null, byte? nivelCuenta = null, bool soloMovimiento = false, bool soloUltimoNivel = false, int tamanoPagina = TamanoAyudaCuenta, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var nivelCuentaTrabajo = nivelCuenta is >= 1 and <= 5 ? nivelCuenta : null;
        var filtro = string.IsNullOrWhiteSpace(textoBusqueda) ? null : textoBusqueda.Trim();
        if (!string.IsNullOrWhiteSpace(filtro) && filtro.Length < 2 && !nivelCuentaTrabajo.HasValue)
        {
            filtro = null;
        }

        var resultado = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            filtro,
            nivelCuentaTrabajo,
            1,
            Math.Clamp(tamanoPagina, 1, TamanoAyudaCuenta),
            soloMovimiento,
            soloUltimoNivel,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idPlanCuenta = x.IdPlanCuenta,
                codigoCuenta = x.CodigoCuenta,
                nombreCuenta = x.NombreCuenta,
                nivelCuenta = x.NivelCuenta,
                requiereCentroCosto = x.RequiereCentroCosto,
                aceptaMovimiento = x.AceptaMovimiento,
                esUltimoNivel = x.EsUltimoNivel
            }),
            total = resultado.TotalRecords,
            limitado = resultado.TotalRecords > resultado.Items.Count
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CargarDefault(CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await planCuentaRepository.CargarDefaultAsync(currentCompanyAccessor.EmpresaId.Value, User.Identity?.Name, cancellationToken);
        TempData["PlanCuentaOk"] = "Plan de cuentas base cargado correctamente.";

        return RedirectToAction(nameof(Index));
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
            modelConError.Monedas = await ObtenerMonedasAsync(cancellationToken);
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
                ColBalance = formulario.ColBalance.Trim().ToUpperInvariant(),
                IdMoneda = formulario.IdMoneda?.Trim().ToUpperInvariant() ?? string.Empty,
                TipoCambio = formulario.TipoCambio?.Trim().ToUpperInvariant() ?? string.Empty,
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
            modelConError.Monedas = await ObtenerMonedasAsync(cancellationToken);
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
        model.Monedas = await ObtenerMonedasAsync(cancellationToken);
        return View("Formulario", model);
    }

    private async Task<List<OpcionCatalogoViewModel>> ObtenerMonedasAsync(CancellationToken cancellationToken)
    {
        var monedas = await monedaRepository.ListarActivasAsync(cancellationToken);

        return monedas
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoMoneda,
                Texto = $"{x.CodigoMoneda} - {x.NombreMoneda}"
            })
            .ToList();
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
                ColBalance = x.ColBalance,
                IdMoneda = x.IdMoneda,
                TipoCambio = x.TipoCambio,
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
                    CuentaPadreTexto = cuentaEditar.IdPlanCuentaPadre.HasValue
                        ? $"{items.FirstOrDefault(x => x.IdPlanCuenta == cuentaEditar.IdPlanCuentaPadre.Value)?.CodigoCuenta ?? string.Empty} - {items.FirstOrDefault(x => x.IdPlanCuenta == cuentaEditar.IdPlanCuentaPadre.Value)?.NombreCuenta ?? string.Empty}".Trim(' ', '-')
                        : string.Empty,
                    CodigoCuenta = cuentaEditar.CodigoCuenta,
                    NombreCuenta = cuentaEditar.NombreCuenta,
                    ColBalance = cuentaEditar.ColBalance,
                    IdMoneda = cuentaEditar.IdMoneda,
                    TipoCambio = cuentaEditar.TipoCambio,
                    AceptaMovimiento = cuentaEditar.AceptaMovimiento,
                    RequiereCentroCosto = cuentaEditar.RequiereCentroCosto,
                    Estado = cuentaEditar.Estado
                }
        };
    }
}
