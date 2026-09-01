using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("PLANCUENTA")]
public class PlanCuentaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPlanCuentaRepository planCuentaRepository,
    IMonedaRepository monedaRepository,
    ICuentaDestinoReglaRepository cuentaDestinoReglaRepository) : Controller
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
        var model = ConstruirViewModel(cuentas.Items, null, null);
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

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("PLANCUENTA", ModulePermissionOperation.Delete)]
    public async Task<IActionResult> Eliminar(int idPlanCuenta, string? textoBusqueda = null, byte? nivelCuenta = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await planCuentaRepository.EliminarAsync(currentCompanyAccessor.EmpresaId.Value, idPlanCuenta, cancellationToken);
            TempData["PlanCuentaOk"] = "Cuenta contable eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PlanCuentaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { textoBusqueda, nivelCuenta, pagina });
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
    [ModulePermission("PLANCUENTA", ModulePermissionOperation.Create)]
    public async Task<IActionResult> CargarDefault(CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await planCuentaRepository.CargarDefaultAsync(currentCompanyAccessor.EmpresaId.Value, User.Identity?.Name, cancellationToken);
            TempData["PlanCuentaOk"] = "Plan de cuentas, parametros, cuentas destino, impuestos y documentos cargados correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PlanCuentaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModuleSavePermission("PLANCUENTA", nameof(PlanCuentaFormViewModel.IdPlanCuenta))]
    public async Task<IActionResult> Guardar(PlanCuentaFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarConfiguracionDestino(formulario.ConfiguracionDestino);
        var guardarConfiguracionDestino = TieneConfiguracionDestino(formulario.ConfiguracionDestino);
        const string campoAnalisis = "Formulario.GeneraDiferenciaPorAnalisis";

        if (guardarConfiguracionDestino)
        {
            if (!formulario.AceptaMovimiento)
            {
                ModelState.AddModelError(string.Empty, "La cuenta debe aceptar movimiento para configurar cuentas destino.");
            }

            ValidarConfiguracionDestino(formulario.ConfiguracionDestino, nameof(formulario.ConfiguracionDestino));
        }

        formulario.IdMoneda = NormalizarMonedaPlanCuenta(formulario.IdMoneda);

        if (formulario.GeneraDiferenciaPorAnalisis)
        {
            if (!formulario.AceptaMovimiento)
            {
                ModelState.AddModelError(campoAnalisis, "La cuenta debe aceptar movimiento para marcarla como analisis.");
            }
        }

        if (!ModelState.IsValid)
        {
            var cuentasConError = await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(cuentasConError, null, null);
            modelConError.Formulario = formulario;
            CompletarTextosFormulario(modelConError.Formulario, cuentasConError);
            modelConError.Monedas = await ObtenerMonedasAsync(cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var cuentaGuardada = await planCuentaRepository.GuardarAsync(new GuardarPlanCuentaRequest
            {
                IdPlanCuenta = formulario.IdPlanCuenta,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdPlanCuentaPadre = formulario.IdPlanCuentaPadre,
                CodigoCuenta = formulario.CodigoCuenta.Trim(),
                NombreCuenta = formulario.NombreCuenta.Trim(),
                ColBalance = formulario.ColBalance.Trim().ToUpperInvariant(),
                IdMoneda = NormalizarMonedaPlanCuenta(formulario.IdMoneda),
                TipoCambio = formulario.TipoCambio?.Trim().ToUpperInvariant() ?? string.Empty,
                AceptaMovimiento = formulario.AceptaMovimiento,
                GeneraDiferenciaPorAnalisis = formulario.GeneraDiferenciaPorAnalisis,
                RequiereCentroCosto = formulario.RequiereCentroCosto,
                Estado = formulario.Estado,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            if (guardarConfiguracionDestino)
            {
                if (!cuentaGuardada.EsUltimoNivel)
                {
                    throw new InvalidOperationException("Solo las cuentas de ultimo nivel pueden configurar cuentas destino.");
                }

                await cuentaDestinoReglaRepository.GuardarAsync(new GuardarCuentaDestinoReglaRequest
                {
                    IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                    IdPlanCuentaOrigen = cuentaGuardada.IdPlanCuenta,
                    Activo = formulario.ConfiguracionDestino.Activo,
                    Observacion = string.IsNullOrWhiteSpace(formulario.ConfiguracionDestino.Observacion)
                        ? null
                        : formulario.ConfiguracionDestino.Observacion.Trim(),
                    UsuarioRegistro = User.Identity?.Name,
                    Detalles = formulario.ConfiguracionDestino.Detalles
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
            }

            TempData["PlanCuentaOk"] = formulario.IdPlanCuenta.HasValue
                ? guardarConfiguracionDestino
                    ? "Cuenta y cuentas destino actualizadas correctamente."
                    : "Cuenta actualizada correctamente."
                : guardarConfiguracionDestino
                    ? "Cuenta y cuentas destino registradas correctamente."
                    : "Cuenta registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var cuentasConError = await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, false, cancellationToken);
            var modelConError = ConstruirViewModel(cuentasConError, null, null);
            modelConError.Formulario = formulario;
            CompletarTextosFormulario(modelConError.Formulario, cuentasConError);
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
        var cuentaDestinoEditar = cuentaEditar is null
            ? null
            : await ObtenerReglaPorCuentaOrigenAsync(currentCompanyAccessor.EmpresaId.Value, cuentaEditar.IdPlanCuenta, cancellationToken);

        var model = ConstruirViewModel(cuentas, cuentaEditar, cuentaDestinoEditar);
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

    private async Task<CuentaDestinoReglaDto?> ObtenerReglaPorCuentaOrigenAsync(int idEmpresa, int idPlanCuentaOrigen, CancellationToken cancellationToken)
    {
        var reglas = await cuentaDestinoReglaRepository.ListarPorEmpresaAsync(idEmpresa, cancellationToken);
        var resumen = reglas.FirstOrDefault(x => x.IdPlanCuentaOrigen == idPlanCuentaOrigen);
        return resumen is null
            ? null
            : await cuentaDestinoReglaRepository.ObtenerAsync(resumen.IdCuentaDestinoRegla, cancellationToken);
    }

    private static void NormalizarConfiguracionDestino(PlanCuentaDestinoConfiguracionViewModel configuracion)
    {
        configuracion.Observacion = string.IsNullOrWhiteSpace(configuracion.Observacion)
            ? null
            : configuracion.Observacion.Trim();

        configuracion.Detalles = configuracion.Detalles
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

    private static bool TieneConfiguracionDestino(PlanCuentaDestinoConfiguracionViewModel configuracion)
    {
        return configuracion.Detalles.Count > 0;
    }

    private static void CompletarTextosFormulario(PlanCuentaFormViewModel formulario, IReadOnlyCollection<PlanCuentaDto> cuentas)
    {
        if (formulario.IdPlanCuentaPadre.HasValue)
        {
            var cuentaPadre = cuentas.FirstOrDefault(x => x.IdPlanCuenta == formulario.IdPlanCuentaPadre.Value);
            formulario.CuentaPadreTexto = cuentaPadre is null
                ? string.Empty
                : $"{cuentaPadre.CodigoCuenta} - {cuentaPadre.NombreCuenta}";
        }

        foreach (var detalle in formulario.ConfiguracionDestino.Detalles)
        {
            if (detalle.IdPlanCuentaDestinoCargo.HasValue)
            {
                var cuentaCargo = cuentas.FirstOrDefault(x => x.IdPlanCuenta == detalle.IdPlanCuentaDestinoCargo.Value);
                detalle.CuentaDestinoCargoTexto = cuentaCargo is null
                    ? detalle.CuentaDestinoCargoTexto
                    : $"{cuentaCargo.CodigoCuenta} - {cuentaCargo.NombreCuenta}";
            }

            if (detalle.IdPlanCuentaDestinoAbono.HasValue)
            {
                var cuentaAbono = cuentas.FirstOrDefault(x => x.IdPlanCuenta == detalle.IdPlanCuentaDestinoAbono.Value);
                detalle.CuentaDestinoAbonoTexto = cuentaAbono is null
                    ? detalle.CuentaDestinoAbonoTexto
                    : $"{cuentaAbono.CodigoCuenta} - {cuentaAbono.NombreCuenta}";
            }
        }
    }

    private void ValidarConfiguracionDestino(PlanCuentaDestinoConfiguracionViewModel configuracion, string prefijo)
    {
        if (configuracion.Detalles.Count == 0)
        {
            ModelState.AddModelError($"{prefijo}.Detalles", "Debe registrar al menos un tramo de cuenta destino.");
            return;
        }

        decimal porcentajeTotal = 0;

        for (var i = 0; i < configuracion.Detalles.Count; i++)
        {
            var detalle = configuracion.Detalles[i];
            var detallePrefijo = $"{prefijo}.Detalles[{i}]";

            if (!detalle.IdPlanCuentaDestinoCargo.HasValue)
            {
                ModelState.AddModelError($"{detallePrefijo}.IdPlanCuentaDestinoCargo", "Seleccione la cuenta destino.");
            }

            if (!detalle.IdPlanCuentaDestinoAbono.HasValue)
            {
                ModelState.AddModelError($"{detallePrefijo}.IdPlanCuentaDestinoAbono", "Seleccione la contrapartida.");
            }

            if (detalle.IdPlanCuentaDestinoCargo.HasValue
                && detalle.IdPlanCuentaDestinoAbono.HasValue
                && detalle.IdPlanCuentaDestinoCargo.Value == detalle.IdPlanCuentaDestinoAbono.Value)
            {
                ModelState.AddModelError($"{detallePrefijo}.IdPlanCuentaDestinoAbono", "Destino y contrapartida no pueden ser la misma cuenta.");
            }

            if (detalle.Activo)
            {
                porcentajeTotal += detalle.Porcentaje;
            }
        }

        if (decimal.Round(porcentajeTotal, 4) != 100)
        {
            ModelState.AddModelError($"{prefijo}.Detalles", "La suma de porcentajes activos debe ser 100.");
        }
    }

    private static PlanCuentaDestinoConfiguracionViewModel CrearConfiguracionDestinoViewModel(CuentaDestinoReglaDto? cuentaDestinoEditar)
    {
        if (cuentaDestinoEditar is null)
        {
            return new PlanCuentaDestinoConfiguracionViewModel();
        }

        return new PlanCuentaDestinoConfiguracionViewModel
        {
            Activo = cuentaDestinoEditar.Activo,
            Observacion = cuentaDestinoEditar.Observacion,
            Detalles = cuentaDestinoEditar.Detalles
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
        };
    }

    private PlanCuentaIndexViewModel ConstruirViewModel(
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        PlanCuentaDto? cuentaEditar,
        CuentaDestinoReglaDto? cuentaDestinoEditar)
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
                IdMoneda = NormalizarMonedaPlanCuenta(x.IdMoneda),
                TipoCambio = x.TipoCambio,
                AceptaMovimiento = x.AceptaMovimiento,
                GeneraDiferenciaPorAnalisis = x.GeneraDiferenciaPorAnalisis,
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
                    IdMoneda = NormalizarMonedaPlanCuenta(cuentaEditar.IdMoneda),
                    TipoCambio = cuentaEditar.TipoCambio,
                    AceptaMovimiento = cuentaEditar.AceptaMovimiento,
                    GeneraDiferenciaPorAnalisis = cuentaEditar.GeneraDiferenciaPorAnalisis,
                    RequiereCentroCosto = cuentaEditar.RequiereCentroCosto,
                    Estado = cuentaEditar.Estado,
                    PermiteConfigurarDestinos = cuentaEditar.EsUltimoNivel || cuentaEditar.AceptaMovimiento || cuentaDestinoEditar is not null,
                    ConfiguracionDestino = CrearConfiguracionDestinoViewModel(cuentaDestinoEditar)
                }
        };
    }

    private static string NormalizarMonedaPlanCuenta(string? idMoneda)
    {
        var valor = (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
        return valor switch
        {
            "S" => "PEN",
            "D" => "USD",
            _ => valor
        };
    }

}
