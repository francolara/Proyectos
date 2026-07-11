using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("TRANSFERENCIAS")]
public class TransferenciaCuentaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPeriodoContableService periodoContableService,
    ICajaBancoRepository cajaBancoRepository,
    ICuentaCorrienteRepository cuentaCorrienteRepository,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ITipoCambioRepository tipoCambioRepository,
    ITipoCambioSyncService tipoCambioSyncService) : Controller
{
    private const int TamanoPagina = 20;

    [HttpGet]
    public async Task<IActionResult> Index(short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var operacionesEmisor = await ObtenerOperacionesAsync("E", cancellationToken);
        var operacionesReceptor = await ObtenerOperacionesAsync("I", cancellationToken);
        var transferencias = await cajaBancoRepository.ListarTransferenciasPaginadoPorEmpresaAsync(
            empresaId,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            pagina,
            TamanoPagina,
            cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            cuentas,
            operacionesEmisor,
            operacionesReceptor,
            transferencias.Items);

        model.TotalTransferencias = transferencias.TotalRecords;
        model.TotalImporteEmisor = transferencias.Items.Sum(x => x.ImporteEmisor);
        model.TotalImporteReceptor = transferencias.Items.Sum(x => x.ImporteReceptor);
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = transferencias.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(empresaId, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["TransferenciaCuentaError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var operacionesEmisor = await ObtenerOperacionesAsync("E", cancellationToken);
        var operacionesReceptor = await ObtenerOperacionesAsync("I", cancellationToken);

        var formulario = new TransferenciaCuentaFormViewModel
        {
            Emisor = new TransferenciaCuentaSeccionFormViewModel
            {
                FechaEmision = new DateOnly(anioTrabajo, mesTrabajo, 1),
                TipoCambio = 1m,
                IdOpeBancaria = operacionesEmisor.FirstOrDefault()?.IdOpeBancaria ?? string.Empty,
                TipoOperacionTexto = operacionesEmisor.FirstOrDefault()?.TipoOperacion ?? string.Empty
            },
            Receptor = new TransferenciaCuentaSeccionFormViewModel
            {
                FechaEmision = new DateOnly(anioTrabajo, mesTrabajo, 1),
                TipoCambio = 1m,
                IdOpeBancaria = operacionesReceptor.FirstOrDefault()?.IdOpeBancaria ?? string.Empty,
                TipoOperacionTexto = operacionesReceptor.FirstOrDefault()?.TipoOperacion ?? string.Empty
            }
        };

        var tipoCambioInicial = await ResolverTipoCambioTransferenciaAsync(empresaId, formulario.Emisor.FechaEmision, cancellationToken);
        if (tipoCambioInicial.HasValue)
        {
            formulario.Emisor.TipoCambio = tipoCambioInicial.Value;
            formulario.Receptor.TipoCambio = tipoCambioInicial.Value;
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            anioTrabajo,
            mesTrabajo,
            null,
            cuentas,
            operacionesEmisor,
            operacionesReceptor,
            []);
        model.Formulario = formulario;
        HidratarFormulario(model);

        return View("Formulario", model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(TransferenciaCuentaFormViewModel formulario, short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        if (await periodoContableService.EstaCerradoAsync(
                empresaId,
                (short)formulario.Emisor.FechaEmision.Year,
                (byte)formulario.Emisor.FechaEmision.Month,
                cancellationToken))
        {
            ModelState.AddModelError(
                string.Empty,
                periodoContableService.ConstruirMensajeBloqueo(
                    (short)formulario.Emisor.FechaEmision.Year,
                    (byte)formulario.Emisor.FechaEmision.Month));
        }

        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var cuentasPorId = cuentas.ToDictionary(x => x.IdBancoConfiguracionEmpresa);
        var operacionesEmisor = await ObtenerOperacionesAsync("E", cancellationToken);
        var operacionesReceptor = await ObtenerOperacionesAsync("I", cancellationToken);
        var operacionesEmisorValidas = operacionesEmisor.Select(x => x.IdOpeBancaria).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var operacionesReceptorValidas = operacionesReceptor.Select(x => x.IdOpeBancaria).ToHashSet(StringComparer.OrdinalIgnoreCase);
        BancoConfiguracionEmpresaDto? cuentaEmisor = null;
        BancoConfiguracionEmpresaDto? cuentaReceptor = null;

        if (formulario.Emisor.IdBancoConfiguracionEmpresa is null || !cuentasPorId.TryGetValue(formulario.Emisor.IdBancoConfiguracionEmpresa.Value, out var cuentaEmisorEncontrada))
        {
            ModelState.AddModelError("Emisor.IdBancoConfiguracionEmpresa", "Seleccione una cuenta corriente emisora activa.");
        }
        else
        {
            cuentaEmisor = cuentaEmisorEncontrada;
            formulario.Emisor.CuentaCorrienteTexto = cuentaEmisor.NroCuentaCorriente;
            formulario.Emisor.MonedaTexto = $"{cuentaEmisor.CodigoMoneda} - {cuentaEmisor.NombreMoneda}";
        }

        if (formulario.Receptor.IdBancoConfiguracionEmpresa is null || !cuentasPorId.TryGetValue(formulario.Receptor.IdBancoConfiguracionEmpresa.Value, out var cuentaReceptorEncontrada))
        {
            ModelState.AddModelError("Receptor.IdBancoConfiguracionEmpresa", "Seleccione una cuenta corriente receptora activa.");
        }
        else
        {
            cuentaReceptor = cuentaReceptorEncontrada;
            formulario.Receptor.CuentaCorrienteTexto = cuentaReceptor.NroCuentaCorriente;
            formulario.Receptor.MonedaTexto = $"{cuentaReceptor.CodigoMoneda} - {cuentaReceptor.NombreMoneda}";
        }

        if (formulario.Emisor.IdBancoConfiguracionEmpresa.HasValue
            && formulario.Receptor.IdBancoConfiguracionEmpresa.HasValue
            && formulario.Emisor.IdBancoConfiguracionEmpresa.Value == formulario.Receptor.IdBancoConfiguracionEmpresa.Value)
        {
            ModelState.AddModelError(string.Empty, "La cuenta corriente emisora debe ser distinta de la receptora.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Emisor.IdOpeBancaria) || !operacionesEmisorValidas.Contains(formulario.Emisor.IdOpeBancaria))
        {
            ModelState.AddModelError("Emisor.IdOpeBancaria", "Seleccione una operacion bancaria de transferencia para el emisor.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Receptor.IdOpeBancaria) || !operacionesReceptorValidas.Contains(formulario.Receptor.IdOpeBancaria))
        {
            ModelState.AddModelError("Receptor.IdOpeBancaria", "Seleccione una operacion bancaria de transferencia para el receptor.");
        }

        var contextoSuscripcion = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(empresaId, cancellationToken);
        if (contextoSuscripcion is null)
        {
            ModelState.AddModelError(string.Empty, "No se pudo resolver la cuenta administradora de la empresa activa.");
        }
        else
        {
            var tipoCambioEmisor = await ResolverTipoCambioAsync(contextoSuscripcion.IdCuentaAdministradora, formulario.Emisor.FechaEmision, cancellationToken);
            if (!tipoCambioEmisor.HasValue)
            {
                ModelState.AddModelError("Emisor.FechaEmision", "No existe un tipo de cambio registrado para la fecha del emisor.");
            }
            else
            {
                formulario.Emisor.TipoCambio = tipoCambioEmisor.Value;
            }

            var tipoCambioReceptor = await ResolverTipoCambioAsync(contextoSuscripcion.IdCuentaAdministradora, formulario.Receptor.FechaEmision, cancellationToken);
            if (!tipoCambioReceptor.HasValue)
            {
                ModelState.AddModelError("Receptor.FechaEmision", "No existe un tipo de cambio registrado para la fecha del receptor.");
            }
            else
            {
                formulario.Receptor.TipoCambio = tipoCambioReceptor.Value;
            }
        }

        if (formulario.Emisor.TipoCambio <= 0 || formulario.Receptor.TipoCambio <= 0)
        {
            ModelState.AddModelError(string.Empty, "El tipo de cambio debe ser mayor a cero en ambas secciones.");
        }

        if (formulario.Emisor.Importe <= 0)
        {
            ModelState.AddModelError("Emisor.Importe", "Ingrese un monto mayor a cero en la seccion emisora.");
        }

        if (cuentaEmisor is not null && cuentaReceptor is not null)
        {
            try
            {
                var importeSugeridoReceptor = CalcularImporteReceptor(
                    cuentaEmisor.CodigoMoneda,
                    cuentaReceptor.CodigoMoneda,
                    formulario.Emisor.Importe,
                    formulario.Emisor.TipoCambio);

                if (string.Equals(cuentaEmisor.CodigoMoneda, cuentaReceptor.CodigoMoneda, StringComparison.OrdinalIgnoreCase))
                {
                    formulario.Receptor.Importe = importeSugeridoReceptor;
                }
                else if (formulario.Receptor.Importe <= 0)
                {
                    formulario.Receptor.Importe = importeSugeridoReceptor;
                }
            }
            catch (InvalidOperationException ex)
            {
                ModelState.AddModelError(string.Empty, ex.Message);
            }
        }

        if (formulario.Receptor.Importe <= 0)
        {
            ModelState.AddModelError("Receptor.Importe", "Ingrese un monto receptor mayor a cero.");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);

        if (!ModelState.IsValid)
        {
            var modelConError = ConstruirViewModel(
                empresaId,
                currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
                anioTrabajo,
                mesTrabajo,
                null,
                cuentas,
                operacionesEmisor,
                operacionesReceptor,
                []);
            modelConError.Formulario = formulario;
            HidratarFormulario(modelConError);
            return View("Formulario", modelConError);
        }

        var resultado = await cajaBancoRepository.GuardarTransferenciaAsync(new GuardarTransferenciaCuentaRequest
        {
            IdEmpresa = empresaId,
            IdBancoConfiguracionEmpresaEmisor = formulario.Emisor.IdBancoConfiguracionEmpresa!.Value,
            IdBancoConfiguracionEmpresaReceptor = formulario.Receptor.IdBancoConfiguracionEmpresa!.Value,
            IdOpeBancariaEmisor = formulario.Emisor.IdOpeBancaria.Trim().ToUpperInvariant(),
            IdOpeBancariaReceptor = formulario.Receptor.IdOpeBancaria.Trim().ToUpperInvariant(),
            FechaEmisionEmisor = formulario.Emisor.FechaEmision,
            FechaEmisionReceptor = formulario.Receptor.FechaEmision,
            TipoCambioEmisor = formulario.Emisor.TipoCambio,
            TipoCambioReceptor = formulario.Receptor.TipoCambio,
            NumeroOperacionEmisor = formulario.Emisor.NumeroOperacion,
            NumeroOperacionReceptor = formulario.Receptor.NumeroOperacion,
            ImporteEmisor = formulario.Emisor.Importe,
            ImporteReceptor = formulario.Receptor.Importe,
            GlosaEmisor = formulario.Emisor.Glosa,
            GlosaReceptor = formulario.Receptor.Glosa,
            ObservacionEmisor = formulario.Emisor.Observacion,
            ObservacionReceptor = formulario.Receptor.Observacion,
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["TransferenciaCuentaOk"] =
            $"Transferencia registrada correctamente. Movimientos {resultado.NumeroMovimientoEmisor}/{resultado.NumeroMovimientoReceptor}. Asientos {resultado.NumeroAsientoEmisor?.ToString() ?? "-"} y {resultado.NumeroAsientoReceptor?.ToString() ?? "-"} .";

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idMovimientoBancoEmisor, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["TransferenciaCuentaError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
        }

        try
        {
            await cajaBancoRepository.EliminarTransferenciaAsync(currentCompanyAccessor.EmpresaId.Value, idMovimientoBancoEmisor, cancellationToken);
            TempData["TransferenciaCuentaOk"] = "Transferencia eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["TransferenciaCuentaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
    }

    private async Task<List<OperacionBancariaDto>> ObtenerOperacionesAsync(string tipoMovimiento, CancellationToken cancellationToken)
    {
        return (await cajaBancoRepository.ListarOperacionesBancariasAsync(tipoMovimiento, null, 100, "T", cancellationToken))
            .OrderBy(x => x.IdOpeBancaria)
            .ThenBy(x => x.TipoOperacion)
            .ToList();
    }

    private static TransferenciaCuentaIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        short anioSeleccionado,
        byte mesSeleccionado,
        string? textoBusqueda,
        IReadOnlyCollection<BancoConfiguracionEmpresaDto> cuentasCorrientes,
        IReadOnlyCollection<OperacionBancariaDto> operacionesEmisor,
        IReadOnlyCollection<OperacionBancariaDto> operacionesReceptor,
        IReadOnlyCollection<TransferenciaCuentaResumenDto> transferencias)
    {
        return new TransferenciaCuentaIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = $"{anioSeleccionado:0000}{mesSeleccionado:00}",
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            AniosDisponibles = Enumerable.Range(anioSeleccionado - 5, 11).ToList(),
            MesesDisponibles = Enumerable.Range(1, 12)
                .Select(x => new MesOpcionViewModel
                {
                    Valor = (byte)x,
                    Nombre = new DateTime(2000, x, 1).ToString("MMMM")
                })
                .ToList(),
            CuentasCorrientesDisponibles = cuentasCorrientes.ToList(),
            OperacionesEmisor = operacionesEmisor.Select(x => new CajaBancoOperacionViewModel
            {
                IdOpeBancaria = x.IdOpeBancaria,
                TipoOperacion = x.TipoOperacion
            }).ToList(),
            OperacionesReceptor = operacionesReceptor.Select(x => new CajaBancoOperacionViewModel
            {
                IdOpeBancaria = x.IdOpeBancaria,
                TipoOperacion = x.TipoOperacion
            }).ToList(),
            Transferencias = transferencias.Select(x => new TransferenciaCuentaResumenItemViewModel
            {
                IdTransferenciaCuenta = x.IdTransferenciaCuenta,
                IdMovimientoBancoEmisor = x.IdMovimientoBancoEmisor,
                IdAsientoEmisor = x.IdAsientoEmisor,
                NumeroMovimientoEmisor = x.NumeroMovimientoEmisor,
                NumeroAsientoEmisor = x.NumeroAsientoEmisor,
                CuentaCorrienteEmisor = x.CuentaCorrienteEmisor,
                MonedaEmisor = x.MonedaEmisor,
                OperacionEmisor = x.OperacionEmisor,
                FechaEmisionEmisor = x.FechaEmisionEmisor,
                NumeroOperacionEmisor = x.NumeroOperacionEmisor,
                ImporteEmisor = x.ImporteEmisor,
                GlosaEmisor = x.GlosaEmisor,
                IdAsientoReceptor = x.IdAsientoReceptor,
                NumeroMovimientoReceptor = x.NumeroMovimientoReceptor,
                NumeroAsientoReceptor = x.NumeroAsientoReceptor,
                CuentaCorrienteReceptor = x.CuentaCorrienteReceptor,
                MonedaReceptor = x.MonedaReceptor,
                OperacionReceptor = x.OperacionReceptor,
                FechaEmisionReceptor = x.FechaEmisionReceptor,
                NumeroOperacionReceptor = x.NumeroOperacionReceptor,
                ImporteReceptor = x.ImporteReceptor,
                GlosaReceptor = x.GlosaReceptor
            }).ToList()
        };
    }

    private static void HidratarFormulario(TransferenciaCuentaIndexViewModel model)
    {
        HidratarSeccion(model.Formulario.Emisor, model.CuentasCorrientesDisponibles, model.OperacionesEmisor);
        HidratarSeccion(model.Formulario.Receptor, model.CuentasCorrientesDisponibles, model.OperacionesReceptor);
    }

    private static void HidratarSeccion(
        TransferenciaCuentaSeccionFormViewModel seccion,
        IReadOnlyCollection<BancoConfiguracionEmpresaDto> cuentas,
        IReadOnlyCollection<CajaBancoOperacionViewModel> operaciones)
    {
        if (seccion.IdBancoConfiguracionEmpresa.HasValue)
        {
            var cuenta = cuentas.FirstOrDefault(x => x.IdBancoConfiguracionEmpresa == seccion.IdBancoConfiguracionEmpresa.Value);
            if (cuenta is not null)
            {
                seccion.CuentaCorrienteTexto = cuenta.NroCuentaCorriente;
                seccion.MonedaTexto = $"{cuenta.CodigoMoneda} - {cuenta.NombreMoneda}";
            }
        }

        if (!string.IsNullOrWhiteSpace(seccion.IdOpeBancaria))
        {
            var operacion = operaciones.FirstOrDefault(x => string.Equals(x.IdOpeBancaria, seccion.IdOpeBancaria, StringComparison.OrdinalIgnoreCase));
            if (operacion is not null)
            {
                seccion.TipoOperacionTexto = operacion.TipoOperacion;
            }
        }
    }

    private static decimal CalcularImporteReceptor(string monedaEmisor, string monedaReceptor, decimal importeEmisor, decimal tipoCambio)
    {
        var monedaEmisorTrabajo = (monedaEmisor ?? string.Empty).Trim().ToUpperInvariant();
        var monedaReceptorTrabajo = (monedaReceptor ?? string.Empty).Trim().ToUpperInvariant();

        if (monedaEmisorTrabajo == monedaReceptorTrabajo)
        {
            return importeEmisor;
        }

        if (monedaEmisorTrabajo == "USD" && monedaReceptorTrabajo == "PEN")
        {
            return Math.Round(importeEmisor * tipoCambio, 2, MidpointRounding.AwayFromZero);
        }

        if (monedaEmisorTrabajo == "PEN" && monedaReceptorTrabajo == "USD")
        {
            return Math.Round(importeEmisor / tipoCambio, 2, MidpointRounding.AwayFromZero);
        }

        throw new InvalidOperationException("Solo se admite conversion automatica entre cuentas en PEN y USD.");
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var hoy = DateTime.Today;
        return (anio ?? (short)hoy.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)hoy.Month);
    }

    private static void NormalizarFormulario(TransferenciaCuentaFormViewModel formulario)
    {
        formulario.Emisor.IdOpeBancaria = (formulario.Emisor.IdOpeBancaria ?? string.Empty).Trim().ToUpperInvariant();
        formulario.Receptor.IdOpeBancaria = (formulario.Receptor.IdOpeBancaria ?? string.Empty).Trim().ToUpperInvariant();
        formulario.Emisor.NumeroOperacion = string.IsNullOrWhiteSpace(formulario.Emisor.NumeroOperacion) ? null : formulario.Emisor.NumeroOperacion.Trim();
        formulario.Receptor.NumeroOperacion = string.IsNullOrWhiteSpace(formulario.Receptor.NumeroOperacion) ? null : formulario.Receptor.NumeroOperacion.Trim();
        formulario.Emisor.Glosa = (formulario.Emisor.Glosa ?? string.Empty).Trim();
        formulario.Receptor.Glosa = (formulario.Receptor.Glosa ?? string.Empty).Trim();
        formulario.Emisor.Observacion = string.IsNullOrWhiteSpace(formulario.Emisor.Observacion) ? null : formulario.Emisor.Observacion.Trim();
        formulario.Receptor.Observacion = string.IsNullOrWhiteSpace(formulario.Receptor.Observacion) ? null : formulario.Receptor.Observacion.Trim();
    }

    private async Task<decimal?> ResolverTipoCambioTransferenciaAsync(int idEmpresa, DateOnly fecha, CancellationToken cancellationToken)
    {
        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(idEmpresa, cancellationToken);
        if (contexto is null)
        {
            return null;
        }

        return await ResolverTipoCambioAsync(contexto.IdCuentaAdministradora, fecha, cancellationToken);
    }

    private async Task<decimal?> ResolverTipoCambioAsync(int idCuentaAdministradora, DateOnly fecha, CancellationToken cancellationToken)
    {
        var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(idCuentaAdministradora, fecha, "USD", cancellationToken);
        if (tipoCambio is null)
        {
            tipoCambio = await tipoCambioSyncService.SincronizarFechaAsync(
                idCuentaAdministradora,
                fecha,
                "USD",
                User.Identity?.Name,
                cancellationToken);
        }

        return tipoCambio?.Venta;
    }
}
