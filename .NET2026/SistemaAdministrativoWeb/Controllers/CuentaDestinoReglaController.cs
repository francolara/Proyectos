using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CUENTASDESTINO")]
public class CuentaDestinoReglaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPlanCuentaRepository planCuentaRepository,
    ICuentaDestinoReglaRepository cuentaDestinoReglaRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
        var reglas = await cuentaDestinoReglaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, textoBusqueda, pagina, TamanoPagina, cancellationToken);
        var model = ConstruirViewModel(
            currentCompanyAccessor.EmpresaId.Value,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            cuentasMovimiento,
            reglas.Items,
            null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalReglas = reglas.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = reglas.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idCuentaDestinoRegla, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idCuentaDestinoRegla, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModuleSavePermission("CUENTASDESTINO", nameof(CuentaDestinoReglaFormViewModel.IdCuentaDestinoRegla))]
    public async Task<IActionResult> Guardar(CuentaDestinoReglaFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);
        ValidarFormulario(formulario);

        if (!ModelState.IsValid)
        {
            var cuentasConError = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
            var reglasConError = await cuentaDestinoReglaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var modelConError = ConstruirViewModel(
                currentCompanyAccessor.EmpresaId.Value,
                currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
                cuentasConError,
                reglasConError,
                null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }

        try
        {
            await cuentaDestinoReglaRepository.GuardarAsync(new GuardarCuentaDestinoReglaRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdPlanCuentaOrigen = formulario.IdPlanCuentaOrigen!.Value,
                Activo = formulario.Activo,
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarCuentaDestinoReglaDetalleRequest
                    {
                        Orden = x.Orden,
                        IdPlanCuentaDestinoCargo = x.IdPlanCuentaDestinoCargo!.Value,
                        IdPlanCuentaDestinoAbono = x.IdPlanCuentaDestinoAbono!.Value,
                        Porcentaje = decimal.Round(x.Porcentaje, 4),
                        Activo = x.Activo
                    })
                    .ToList()
            }, cancellationToken);

            TempData["CuentaDestinoOk"] = "Cuenta destino guardada correctamente.";
            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var cuentasConError = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
            var reglasConError = await cuentaDestinoReglaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var modelConError = ConstruirViewModel(
                currentCompanyAccessor.EmpresaId.Value,
                currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
                cuentasConError,
                reglasConError,
                null);
            modelConError.Formulario = formulario;
            return View("Formulario", modelConError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idCuentaDestinoRegla, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await cuentaDestinoReglaRepository.EliminarAsync(idCuentaDestinoRegla, cancellationToken);
        TempData["CuentaDestinoOk"] = "Cuenta destino eliminada correctamente.";
        return RedirectToAction(nameof(Index));
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idCuentaDestinoRegla, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
        var reglas = await cuentaDestinoReglaRepository.ListarPorEmpresaAsync(empresaId, cancellationToken);
        var reglaEditar = idCuentaDestinoRegla.HasValue
            ? await cuentaDestinoReglaRepository.ObtenerAsync(idCuentaDestinoRegla.Value, cancellationToken)
            : null;

        if (reglaEditar is not null && reglaEditar.IdEmpresa != empresaId)
        {
            reglaEditar = null;
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            cuentasMovimiento,
            reglas,
            reglaEditar);

        return View("Formulario", model);
    }

    private static void NormalizarFormulario(CuentaDestinoReglaFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuentaDestinoCargo.HasValue
                     || x.IdPlanCuentaDestinoAbono.HasValue
                     || x.Porcentaje > 0
                     || x.Activo)
            .Select((x, index) =>
            {
                x.Orden = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(CuentaDestinoReglaFormViewModel formulario)
    {
        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos un tramo.");
            return;
        }

        decimal porcentajeTotal = 0;

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Detalles[{i}]";

            if (!detalle.IdPlanCuentaDestinoCargo.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuentaDestinoCargo", "Seleccione la cuenta cargo.");
            }

            if (!detalle.IdPlanCuentaDestinoAbono.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuentaDestinoAbono", "Seleccione la cuenta abono.");
            }

            if (detalle.IdPlanCuentaDestinoCargo.HasValue
                && detalle.IdPlanCuentaDestinoAbono.HasValue
                && detalle.IdPlanCuentaDestinoCargo.Value == detalle.IdPlanCuentaDestinoAbono.Value)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuentaDestinoAbono", "Cargo y abono no pueden ser la misma cuenta.");
            }

            if (detalle.Activo)
            {
                porcentajeTotal += detalle.Porcentaje;
            }
        }

        if (decimal.Round(porcentajeTotal, 4) != 100)
        {
            ModelState.AddModelError(string.Empty, "La suma de porcentajes activos debe ser 100.");
        }
    }

    private static CuentaDestinoReglaIndexViewModel ConstruirViewModel(
        int idEmpresa,
        string empresaNombre,
        IReadOnlyCollection<PlanCuentaDto> cuentasMovimiento,
        IReadOnlyCollection<CuentaDestinoReglaResumenDto> reglas,
        CuentaDestinoReglaDto? reglaEditar)
    {
        var reglasItems = reglas
            .Select(x => new CuentaDestinoReglaResumenItemViewModel
            {
                IdCuentaDestinoRegla = x.IdCuentaDestinoRegla,
                IdPlanCuentaOrigen = x.IdPlanCuentaOrigen,
                CodigoCuentaOrigen = x.CodigoCuentaOrigen,
                NombreCuentaOrigen = x.NombreCuentaOrigen,
                Activo = x.Activo,
                Observacion = x.Observacion,
                CantidadTramos = x.CantidadTramos,
                PorcentajeTotal = x.PorcentajeTotal
            })
            .OrderBy(x => x.CodigoCuentaOrigen)
            .ToList();

        return new CuentaDestinoReglaIndexViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = empresaNombre,
            TotalReglas = reglasItems.Count,
            TotalActivas = reglasItems.Count(x => x.Activo),
            TotalTramos = reglasItems.Sum(x => x.CantidadTramos),
            PorcentajeConfigurado = reglasItems.Sum(x => x.PorcentajeTotal),
            CuentasMovimiento = cuentasMovimiento
                .Where(x => x.Estado)
                .OrderBy(x => x.CodigoCuenta)
                .ToList(),
            Reglas = reglasItems,
            Formulario = reglaEditar is null
                ? new CuentaDestinoReglaFormViewModel()
                : new CuentaDestinoReglaFormViewModel
                {
                    IdCuentaDestinoRegla = reglaEditar.IdCuentaDestinoRegla,
                    IdPlanCuentaOrigen = reglaEditar.IdPlanCuentaOrigen,
                    CuentaOrigenTexto = $"{reglaEditar.CodigoCuentaOrigen} - {reglaEditar.NombreCuentaOrigen}",
                    Observacion = reglaEditar.Observacion,
                    Activo = reglaEditar.Activo,
                    Detalles = reglaEditar.Detalles
                        .OrderBy(x => x.Orden)
                        .Select(x => new CuentaDestinoReglaDetalleFormViewModel
                        {
                            IdCuentaDestinoReglaDetalle = x.IdCuentaDestinoReglaDetalle,
                            Orden = x.Orden,
                            IdPlanCuentaDestinoCargo = x.IdPlanCuentaDestinoCargo,
                            CuentaDestinoCargoTexto = $"{x.CodigoCuentaDestinoCargo} - {x.NombreCuentaDestinoCargo}",
                            IdPlanCuentaDestinoAbono = x.IdPlanCuentaDestinoAbono,
                            CuentaDestinoAbonoTexto = $"{x.CodigoCuentaDestinoAbono} - {x.NombreCuentaDestinoAbono}",
                            Porcentaje = x.Porcentaje,
                            Activo = x.Activo
                        })
                        .ToList()
                }
        };
    }
}
