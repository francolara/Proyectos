using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class ConfiguracionContabilizacionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IConfiguracionContabilizacionRepository configuracionRepository,
    IOrigenRepository origenRepository,
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

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var configuraciones = await configuracionRepository.ListarPaginadoPorEmpresaAsync(empresaId, textoBusqueda, pagina, TamanoPagina, cancellationToken);
        var origenes = await origenRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken);
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken);
        var model = ConstruirViewModel(configuraciones.Items, origenes, cuentas, null);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TotalConfiguraciones = configuraciones.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = configuraciones.TotalRecords
        };
        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idConfiguracionContabilizacion, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(ConfiguracionContabilizacionFormViewModel formulario, CancellationToken cancellationToken)
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
            var empresaIdError = currentCompanyAccessor.EmpresaId.Value;
            var configuracionesError = await configuracionRepository.ListarPorEmpresaAsync(empresaIdError, cancellationToken);
            var origenesError = await origenRepository.ListarPorEmpresaAsync(empresaIdError, false, cancellationToken);
            var cuentasError = await planCuentaRepository.ListarPorEmpresaAsync(empresaIdError, true, cancellationToken);
            var modelError = ConstruirViewModel(configuracionesError, origenesError, cuentasError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }

        try
        {
            await configuracionRepository.GuardarAsync(new GuardarConfiguracionContabilizacionRequest
            {
                IdConfiguracionContabilizacion = formulario.IdConfiguracionContabilizacion,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                ModuloOperacion = formulario.ModuloOperacion.Trim().ToUpperInvariant(),
                EscenarioOperacion = formulario.EscenarioOperacion.Trim().ToUpperInvariant(),
                IdOrigen = formulario.IdOrigen!.Value,
                Descripcion = formulario.Descripcion.Trim(),
                GeneraAsientoAutomatico = formulario.GeneraAsientoAutomatico,
                UsaTipoCambio = formulario.UsaTipoCambio,
                Activo = formulario.Activo,
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarConfiguracionContabilizacionDetalleRequest
                    {
                        Orden = x.Orden,
                        ComponenteContable = x.ComponenteContable.Trim().ToUpperInvariant(),
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        NaturalezaMovimiento = x.NaturalezaMovimiento.Trim().ToUpperInvariant(),
                        Activo = x.Activo
                    })
                    .ToList()
            }, cancellationToken);

            TempData["ConfiguracionContabilizacionOk"] = formulario.IdConfiguracionContabilizacion.HasValue
                ? "Configuracion contable actualizada correctamente."
                : "Configuracion contable registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var empresaIdError = currentCompanyAccessor.EmpresaId.Value;
            var configuracionesError = await configuracionRepository.ListarPorEmpresaAsync(empresaIdError, cancellationToken);
            var origenesError = await origenRepository.ListarPorEmpresaAsync(empresaIdError, false, cancellationToken);
            var cuentasError = await planCuentaRepository.ListarPorEmpresaAsync(empresaIdError, true, cancellationToken);
            var modelError = ConstruirViewModel(configuracionesError, origenesError, cuentasError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idConfiguracionContabilizacion, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await configuracionRepository.EliminarAsync(idConfiguracionContabilizacion, cancellationToken);
        TempData["ConfiguracionContabilizacionOk"] = "Configuracion contable eliminada correctamente.";
        return RedirectToAction(nameof(Index));
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idConfiguracionContabilizacion, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var configuraciones = await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken);
        var origenes = await origenRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken);
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken);
        var configuracionEditar = idConfiguracionContabilizacion.HasValue
            ? await configuracionRepository.ObtenerAsync(idConfiguracionContabilizacion.Value, cancellationToken)
            : null;

        if (configuracionEditar is not null && configuracionEditar.IdEmpresa != empresaId)
        {
            configuracionEditar = null;
        }

        return View("Formulario", ConstruirViewModel(configuraciones, origenes, cuentas, configuracionEditar));
    }

    private static void NormalizarFormulario(ConfiguracionContabilizacionFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue || !string.IsNullOrWhiteSpace(x.ComponenteContable))
            .Select((x, index) =>
            {
                x.Orden = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(ConfiguracionContabilizacionFormViewModel formulario)
    {
        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos un componente contable.");
            return;
        }

        var componentesActivos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (!detalle.IdPlanCuenta.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione una cuenta.");
            }

            if (detalle.Activo && !componentesActivos.Add(detalle.ComponenteContable))
            {
                ModelState.AddModelError($"{prefijo}.ComponenteContable", "No repita el mismo componente activo.");
            }
        }
    }

    private ConfiguracionContabilizacionIndexViewModel ConstruirViewModel(
        IReadOnlyCollection<ConfiguracionContabilizacionResumenDto> configuraciones,
        IReadOnlyCollection<OrigenDto> origenes,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        ConfiguracionContabilizacionDto? configuracionEditar)
    {
        var items = configuraciones
            .Select(x => new ConfiguracionContabilizacionResumenItemViewModel
            {
                IdConfiguracionContabilizacion = x.IdConfiguracionContabilizacion,
                ModuloOperacion = x.ModuloOperacion,
                EscenarioOperacion = x.EscenarioOperacion,
                CodigoOrigen = x.CodigoOrigen,
                NombreOrigen = x.NombreOrigen,
                Descripcion = x.Descripcion,
                GeneraAsientoAutomatico = x.GeneraAsientoAutomatico,
                UsaTipoCambio = x.UsaTipoCambio,
                Activo = x.Activo,
                CantidadComponentes = x.CantidadComponentes
            })
            .OrderBy(x => x.ModuloOperacion)
            .ThenBy(x => x.EscenarioOperacion)
            .ToList();

        return new ConfiguracionContabilizacionIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalConfiguraciones = items.Count,
            TotalActivas = items.Count(x => x.Activo),
            TotalAutomaticas = items.Count(x => x.GeneraAsientoAutomatico),
            Configuraciones = items,
            Origenes = origenes.Where(x => x.Estado).OrderBy(x => x.CodigoOrigen).ToList(),
            CuentasMovimiento = cuentas.Where(x => x.Estado).OrderBy(x => x.CodigoCuenta).ToList(),
            Formulario = configuracionEditar is null
                ? new ConfiguracionContabilizacionFormViewModel
                {
                    IdOrigen = origenes.FirstOrDefault(x => x.Estado)?.IdOrigen
                }
                : new ConfiguracionContabilizacionFormViewModel
                {
                    IdConfiguracionContabilizacion = configuracionEditar.IdConfiguracionContabilizacion,
                    ModuloOperacion = configuracionEditar.ModuloOperacion,
                    EscenarioOperacion = configuracionEditar.EscenarioOperacion,
                    IdOrigen = configuracionEditar.IdOrigen,
                    Descripcion = configuracionEditar.Descripcion,
                    GeneraAsientoAutomatico = configuracionEditar.GeneraAsientoAutomatico,
                    UsaTipoCambio = configuracionEditar.UsaTipoCambio,
                    Activo = configuracionEditar.Activo,
                    Detalles = configuracionEditar.Detalles
                        .OrderBy(x => x.Orden)
                        .Select(x => new ConfiguracionContabilizacionDetalleFormViewModel
                        {
                            Orden = x.Orden,
                            ComponenteContable = x.ComponenteContable,
                            IdPlanCuenta = x.IdPlanCuenta,
                            NaturalezaMovimiento = x.NaturalezaMovimiento,
                            Activo = x.Activo
                        })
                        .ToList()
                }
        };
    }
}
