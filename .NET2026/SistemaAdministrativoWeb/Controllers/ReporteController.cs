using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;
using System.Security.Claims;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("REPORTES")]
public class ReporteController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IEmpresaRepository empresaRepository,
    IPlanCuentaRepository planCuentaRepository,
    IAnalisisCuentaRepository analisisCuentaRepository,
    IBalanceComprobacionRepository balanceComprobacionRepository,
    IRegistroVentasRepository registroVentasRepository,
    IRegistroComprasRepository registroComprasRepository,
    IClienteRepository clienteRepository,
    IProveedorRepository proveedorRepository,
    ILibroDiarioRepository libroDiarioRepository,
    ILibroMayorRepository libroMayorRepository,
    IAsientoRepository asientoRepository) : Controller
{
    [HttpGet]
    public async Task<IActionResult> VoucherContable(int idAsiento, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var asiento = await asientoRepository.ObtenerAsync(idAsiento, cancellationToken);
        if (asiento is null || asiento.IdEmpresa != currentCompanyAccessor.EmpresaId.Value)
        {
            return NotFound();
        }

        var rucEmpresa = string.Empty;
        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (!string.IsNullOrWhiteSpace(aspNetUserId))
        {
            var empresas = await empresaRepository.ListarPorUsuarioAsync(aspNetUserId, cancellationToken);
            rucEmpresa = empresas.FirstOrDefault(x => x.IdEmpresa == currentCompanyAccessor.EmpresaId.Value)?.Ruc ?? string.Empty;
        }

        var model = new VoucherContableViewModel
        {
            IdEmpresa = asiento.IdEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            RucEmpresa = rucEmpresa,
            IdAsiento = asiento.IdAsiento,
            Periodo = asiento.Periodo,
            NumeroAsiento = asiento.NumeroAsiento,
            CodigoOrigen = asiento.CodigoOrigen,
            NombreOrigen = asiento.NombreOrigen,
            Glosa = asiento.Glosa,
            Moneda = asiento.CodigoMoneda,
            TipoCambio = asiento.TipoCambio,
            FechaEmision = asiento.FechaEmision,
            ReferenciaExterna = asiento.ReferenciaExterna ?? string.Empty,
            Observacion = asiento.Observacion ?? string.Empty,
            MuestraColumnaDolares = asiento.Detalles.Any(x => x.TotalImporteD != 0m)
        };

        model.Detalles = asiento.Detalles
            .OrderBy(x => x.Item)
            .Select(x => new VoucherContableItemViewModel
            {
                Item = x.Item,
                CodigoCuenta = x.CodigoCuenta,
                NombreCuenta = x.NombreCuenta,
                GlosaDetalle = x.GlosaDetalle ?? string.Empty,
                NumeroDocumento = x.NumeroDocumento ?? string.Empty,
                TipoDocumento = x.TipoDocumento ?? string.Empty,
                Serie = x.Serie ?? string.Empty,
                Referencia = x.ReferenciaLinea ?? string.Empty,
                CentroCosto = x.CodigoCentroCosto ?? string.Empty,
                TipoCambio = x.TipoCambioLinea ?? asiento.TipoCambio,
                Debe = x.Debe,
                Haber = x.Haber,
                ImporteSoles = x.TotalImporteS,
                ImporteDolares = x.TotalImporteD
            })
            .ToList();

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> AnalisisCuentas(
        short? anio = null,
        byte? mes = null,
        string? cuentaDesde = null,
        string? cuentaHasta = null,
        string? auxiliar = null,
        string? moneda = null,
        string? estado = null,
        string? tipo = null,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;

        var model = new AnalisisCuentaViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            PeriodoConsulta = $"{anioTrabajo:0000}{mesTrabajo:00}",
            CuentaDesde = cuentaDesde?.Trim() ?? string.Empty,
            CuentaHasta = cuentaHasta?.Trim() ?? string.Empty,
            Auxiliar = auxiliar?.Trim() ?? string.Empty,
            MonedaSeleccionada = NormalizarMoneda(moneda),
            EstadoSeleccionado = NormalizarEstado(estado),
            TipoSeleccionado = NormalizarTipo(tipo),
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            MesesDisponibles = ListarMesesCalendario(),
            EstadosDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = "T", Texto = "Todos" },
                new OpcionCatalogoViewModel { Valor = "P", Texto = "Pendientes" },
                new OpcionCatalogoViewModel { Valor = "C", Texto = "Cancelados" }
            ],
            MonedasDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = "PEN", Texto = "Soles" },
                new OpcionCatalogoViewModel { Valor = "USD", Texto = "Dolares" }
            ],
            TiposDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = "0", Texto = "Detallado" },
                new OpcionCatalogoViewModel { Valor = "1", Texto = "Auxiliar y documento" },
                new OpcionCatalogoViewModel { Valor = "2", Texto = "Por auxiliar" }
            ]
        };

        if (!consultar)
        {
            return View(model);
        }

        if (EsRangoCuentaInvalido(model.CuentaDesde, model.CuentaHasta))
        {
            model.MensajeError = "Cuenta hasta debe ser mayor o igual que Cuenta desde.";
            return View(model);
        }

        try
        {
            var resultados = await analisisCuentaRepository.ListarAsync(new AnalisisCuentaRequest
            {
                IdEmpresa = idEmpresa,
                Periodo = model.PeriodoConsulta,
                CuentaDesde = model.CuentaDesde,
                CuentaHasta = model.CuentaHasta,
                Auxiliar = model.Auxiliar,
                Moneda = model.MonedaSeleccionada,
                Estado = model.EstadoSeleccionado,
                Tipo = model.TipoSeleccionado
            }, cancellationToken);

            model.Resultados = resultados
                .Select(x => new AnalisisCuentaItemViewModel
                {
                    CodigoCuenta = x.CodigoCuenta,
                    NombreCuenta = x.NombreCuenta,
                    Auxiliar = x.Auxiliar,
                    NombreAuxiliar = x.NombreAuxiliar,
                    TipoDocumento = x.TipoDocumento,
                    Serie = x.Serie,
                    NumeroReferencia = x.NumeroReferencia,
                    Periodo = x.Periodo,
                    Comprobante = x.Comprobante,
                    GlosaDetalle = x.GlosaDetalle,
                    FechaEmision = x.FechaEmision,
                    TipoCambio = x.TipoCambio,
                    Debe = x.Debe,
                    Haber = x.Haber,
                    DebeDolares = x.DebeDolares,
                    HaberDolares = x.HaberDolares
                })
                .ToList();

            model.TotalDebe = model.Resultados.Sum(x => x.Debe);
            model.TotalHaber = model.Resultados.Sum(x => x.Haber);
            model.TotalDebeDolares = model.Resultados.Sum(x => x.DebeDolares);
            model.TotalHaberDolares = model.Resultados.Sum(x => x.HaberDolares);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> BalanceComprobacion(
        short? anio = null,
        byte? periodoDesde = null,
        byte? periodoHasta = null,
        string? moneda = null,
        byte? grado = null,
        bool todasLasCuentas = true,
        string? cuentaDesde = null,
        string? cuentaHasta = null,
        bool filtrarGrado = true,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, periodoDesdeTrabajo, periodoHastaTrabajo) = NormalizarRangoPeriodosContables(anio, periodoDesde, periodoHasta);
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(idEmpresa, true, cancellationToken);
        var gradosMaximos = cuentas.Count == 0 ? 1 : Math.Clamp(cuentas.Max(x => x.NivelCuenta), 1, 9);

        var model = new BalanceComprobacionViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            PeriodoDesdeSeleccionado = periodoDesdeTrabajo,
            PeriodoHastaSeleccionado = periodoHastaTrabajo,
            MonedaSeleccionada = NormalizarMoneda(moneda),
            GradoSeleccionado = NormalizarGrado(grado, gradosMaximos),
            TodasLasCuentas = todasLasCuentas,
            CuentaDesde = todasLasCuentas ? string.Empty : cuentaDesde?.Trim() ?? string.Empty,
            CuentaHasta = todasLasCuentas ? string.Empty : cuentaHasta?.Trim() ?? string.Empty,
            FiltrarGrado = filtrarGrado,
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            PeriodosDisponibles = ListarPeriodosContables(),
            MonedasDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = "PEN", Texto = "Soles" },
                new OpcionCatalogoViewModel { Valor = "USD", Texto = "Dolares" }
            ],
            GradosDisponibles = Enumerable.Range(1, gradosMaximos)
                .Select(x => new OpcionCatalogoViewModel
                {
                    Valor = x.ToString(),
                    Texto = $"Grado {x}"
                })
                .ToList()
        };

        if (!consultar)
        {
            return View(model);
        }

        if (!model.TodasLasCuentas && EsRangoCuentaInvalido(model.CuentaDesde, model.CuentaHasta))
        {
            model.MensajeError = "Cuenta hasta debe ser mayor o igual que Cuenta desde.";
            return View(model);
        }

        try
        {
            var resultados = await balanceComprobacionRepository.ListarAsync(new BalanceComprobacionRequest
            {
                IdEmpresa = idEmpresa,
                Anio = model.AnioSeleccionado,
                PeriodoDesde = model.PeriodoDesdeSeleccionado,
                PeriodoHasta = model.PeriodoHastaSeleccionado,
                Moneda = model.MonedaSeleccionada,
                Grado = model.GradoSeleccionado,
                TodasLasCuentas = model.TodasLasCuentas,
                CuentaDesde = model.CuentaDesde,
                CuentaHasta = model.CuentaHasta,
                FiltrarGrado = model.FiltrarGrado
            }, cancellationToken);

            var resultadosFiltrados = resultados
                .Where(x => model.FiltrarGrado
                    ? x.GradoCuenta == model.GradoSeleccionado
                    : x.GradoCuenta <= model.GradoSeleccionado)
                .ToList();

            model.Resultados = resultadosFiltrados
                .Select(x =>
                {
                    var diferencia = x.Debe - x.Haber;
                    var resultadoDebe = diferencia > 0 ? diferencia : 0m;
                    var resultadoHaber = diferencia < 0 ? diferencia * -1 : 0m;
                    var colBalance = (x.ColBalance ?? string.Empty).Trim().ToUpperInvariant();

                    return new BalanceComprobacionItemViewModel
                    {
                        CodigoCuenta = x.CodigoCuenta,
                        NombreCuenta = x.NombreCuenta,
                        ColBalance = colBalance,
                        GradoCuenta = x.GradoCuenta,
                        DebAnt = x.DebAnt,
                        HabAnt = x.HabAnt,
                        DebMes = x.DebMes,
                        HabMes = x.HabMes,
                        Debe = x.Debe,
                        Haber = x.Haber,
                        ResultadoDebe = resultadoDebe,
                        ResultadoHaber = resultadoHaber,
                        Activo = colBalance == "I" ? resultadoDebe : 0m,
                        Pasivo = colBalance == "I" ? resultadoHaber : 0m,
                        PerdidaNaturaleza = colBalance is "N" or "R" ? resultadoDebe : 0m,
                        GananciaNaturaleza = colBalance is "N" or "R" ? resultadoHaber : 0m,
                        PerdidaFuncion = colBalance is "F" or "R" ? resultadoDebe : 0m,
                        GananciaFuncion = colBalance is "F" or "R" ? resultadoHaber : 0m
                    };
                })
                .ToList();

            model.TotalDebAnt = model.Resultados.Sum(x => x.DebAnt);
            model.TotalHabAnt = model.Resultados.Sum(x => x.HabAnt);
            model.TotalDebMes = model.Resultados.Sum(x => x.DebMes);
            model.TotalHabMes = model.Resultados.Sum(x => x.HabMes);
            model.TotalDebe = model.Resultados.Sum(x => x.Debe);
            model.TotalHaber = model.Resultados.Sum(x => x.Haber);
            model.TotalResultadoDebe = model.Resultados.Sum(x => x.ResultadoDebe);
            model.TotalResultadoHaber = model.Resultados.Sum(x => x.ResultadoHaber);
            model.TotalActivo = model.Resultados.Sum(x => x.Activo);
            model.TotalPasivo = model.Resultados.Sum(x => x.Pasivo);
            model.TotalPerdidaNaturaleza = model.Resultados.Sum(x => x.PerdidaNaturaleza);
            model.TotalGananciaNaturaleza = model.Resultados.Sum(x => x.GananciaNaturaleza);
            model.TotalPerdidaFuncion = model.Resultados.Sum(x => x.PerdidaFuncion);
            model.TotalGananciaFuncion = model.Resultados.Sum(x => x.GananciaFuncion);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> RegistroVentas(
        short? anio = null,
        byte? mes = null,
        string? codigoPersona = null,
        string? numeroDocumento = null,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;

        var model = new RegistroVentasViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            CodigoPersona = codigoPersona?.Trim() ?? string.Empty,
            NumeroComprobante = numeroDocumento?.Trim() ?? string.Empty,
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };

        if (!consultar)
        {
            return View(model);
        }

        try
        {
            var resultados = await registroVentasRepository.ListarAsync(new RegistroVentasRequest
            {
                IdEmpresa = idEmpresa,
                Anio = model.AnioSeleccionado,
                Mes = model.MesSeleccionado,
                CodigoPersona = model.CodigoPersona,
                NumeroComprobante = model.NumeroComprobante
            }, cancellationToken);

            model.Resultados = resultados
                .Select(x => new RegistroVentasItemViewModel
                {
                    FechaEmision = x.FechaEmision,
                    FechaContabilizacion = x.FechaContabilizacion,
                    TipoComprobante = x.TipoComprobante,
                    DescripcionTipoComprobante = x.DescripcionTipoComprobante,
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoPersona = x.CodigoPersona,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    NombrePersona = x.NombrePersona,
                    CodigoMoneda = x.CodigoMoneda,
                    TipoCambio = x.TipoCambio,
                    BaseImponible = x.BaseImponible,
                    Descuento = x.Descuento,
                    TotalExonerado = x.TotalExonerado,
                    TotalInafecto = x.TotalInafecto,
                    Igv = x.Igv,
                    Isc = x.Isc,
                    OtrosTributos = x.OtrosTributos,
                    Icbper = x.Icbper,
                    Redondeo = x.Redondeo,
                    ImporteTotal = x.ImporteTotal,
                    Estado = x.Estado,
                    Observacion = x.Observacion
                })
                .ToList();

            model.TotalBaseImponible = model.Resultados.Sum(x => x.BaseImponible);
            model.TotalDescuento = model.Resultados.Sum(x => x.Descuento);
            model.TotalExonerado = model.Resultados.Sum(x => x.TotalExonerado);
            model.TotalInafecto = model.Resultados.Sum(x => x.TotalInafecto);
            model.TotalIgv = model.Resultados.Sum(x => x.Igv);
            model.TotalIsc = model.Resultados.Sum(x => x.Isc);
            model.TotalOtrosTributos = model.Resultados.Sum(x => x.OtrosTributos);
            model.TotalIcbper = model.Resultados.Sum(x => x.Icbper);
            model.TotalRedondeo = model.Resultados.Sum(x => x.Redondeo);
            model.TotalImporte = model.Resultados.Sum(x => x.ImporteTotal);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> RegistroCompras(
        short? anio = null,
        byte? mes = null,
        string? codigoPersona = null,
        string? numeroDocumento = null,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;

        var model = new RegistroComprasViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            CodigoPersona = codigoPersona?.Trim() ?? string.Empty,
            NumeroComprobante = numeroDocumento?.Trim() ?? string.Empty,
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };

        if (!consultar)
        {
            return View(model);
        }

        try
        {
            var resultados = await registroComprasRepository.ListarAsync(new RegistroComprasRequest
            {
                IdEmpresa = idEmpresa,
                Anio = model.AnioSeleccionado,
                Mes = model.MesSeleccionado,
                CodigoPersona = model.CodigoPersona,
                NumeroComprobante = model.NumeroComprobante
            }, cancellationToken);

            model.Resultados = resultados
                .Select(x => new RegistroComprasItemViewModel
                {
                    FechaEmision = x.FechaEmision,
                    FechaContabilizacion = x.FechaContabilizacion,
                    TipoComprobante = x.TipoComprobante,
                    DescripcionTipoComprobante = x.DescripcionTipoComprobante,
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoPersona = x.CodigoPersona,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    NombrePersona = x.NombrePersona,
                    CodigoMoneda = x.CodigoMoneda,
                    TipoCambio = x.TipoCambio,
                    BaseImponibleGravada = x.BaseImponibleGravada,
                    IgvGravado = x.IgvGravado,
                    BaseImponibleGasto = x.BaseImponibleGasto,
                    IgvGasto = x.IgvGasto,
                    BaseImponibleSinCredito = x.BaseImponibleSinCredito,
                    IgvSinCredito = x.IgvSinCredito,
                    TotalExonerado = x.TotalExonerado,
                    TotalInafecto = x.TotalInafecto,
                    OtrosTributos = x.OtrosTributos,
                    Icbper = x.Icbper,
                    Retencion = x.Retencion,
                    ImporteDetraccion = x.ImporteDetraccion,
                    ImportePercepcion = x.ImportePercepcion,
                    ImporteTotal = x.ImporteTotal,
                    Estado = x.Estado,
                    Observacion = x.Observacion
                })
                .ToList();

            model.TotalBaseImponibleGravada = model.Resultados.Sum(x => x.BaseImponibleGravada);
            model.TotalIgvGravado = model.Resultados.Sum(x => x.IgvGravado);
            model.TotalBaseImponibleGasto = model.Resultados.Sum(x => x.BaseImponibleGasto);
            model.TotalIgvGasto = model.Resultados.Sum(x => x.IgvGasto);
            model.TotalBaseImponibleSinCredito = model.Resultados.Sum(x => x.BaseImponibleSinCredito);
            model.TotalIgvSinCredito = model.Resultados.Sum(x => x.IgvSinCredito);
            model.TotalExonerado = model.Resultados.Sum(x => x.TotalExonerado);
            model.TotalInafecto = model.Resultados.Sum(x => x.TotalInafecto);
            model.TotalOtrosTributos = model.Resultados.Sum(x => x.OtrosTributos);
            model.TotalIcbper = model.Resultados.Sum(x => x.Icbper);
            model.TotalRetencion = model.Resultados.Sum(x => x.Retencion);
            model.TotalDetraccion = model.Resultados.Sum(x => x.ImporteDetraccion);
            model.TotalPercepcion = model.Resultados.Sum(x => x.ImportePercepcion);
            model.TotalImporte = model.Resultados.Sum(x => x.ImporteTotal);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> BuscarPersonasRegistroAyuda(
        string? tipoRegistro = null,
        string? buscar = null,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var tipoTrabajo = (tipoRegistro ?? string.Empty).Trim().ToUpperInvariant();
        var criterio = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;

        if (tipoTrabajo == "VEN")
        {
            var clientes = await clienteRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
            var items = clientes
                .Where(x => criterio is null
                    || x.CodigoCliente.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                    || x.NumeroDocumento.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                    || x.NombreCompleto.Contains(criterio, StringComparison.OrdinalIgnoreCase))
                .OrderBy(x => x.NombreCompleto)
                .Take(30)
                .Select(x => new
                {
                    codigoPersona = x.NumeroDocumento,
                    numeroDocumentoPersona = x.NumeroDocumento,
                    nombreCompleto = x.NombreCompleto
                })
                .ToList();

            return Json(new { ok = true, items });
        }

        if (tipoTrabajo == "COM")
        {
            var proveedores = await proveedorRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
            var items = proveedores
                .Where(x => criterio is null
                    || x.CodigoProveedor.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                    || x.NumeroDocumento.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                    || x.NombreCompleto.Contains(criterio, StringComparison.OrdinalIgnoreCase))
                .OrderBy(x => x.NombreCompleto)
                .Take(30)
                .Select(x => new
                {
                    codigoPersona = x.NumeroDocumento,
                    numeroDocumentoPersona = x.NumeroDocumento,
                    nombreCompleto = x.NombreCompleto
                })
                .ToList();

            return Json(new { ok = true, items });
        }

        return BadRequest(new { ok = false, mensaje = "El tipo de registro solicitado no es valido." });
    }

    [HttpGet]
    public async Task<IActionResult> LibroDiario(
        short? anio = null,
        byte? periodo = null,
        string? modo = null,
        string? cuentaDesde = null,
        string? cuentaHasta = null,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, periodoTrabajo) = NormalizarPeriodoContable(anio, periodo);

        var model = new LibroDiarioViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            PeriodoSeleccionado = periodoTrabajo,
            PeriodoConsulta = $"{anioTrabajo:0000}{periodoTrabajo:00}",
            MonedaSeleccionada = "PEN",
            ModoSeleccionado = NormalizarModoLibroDiario(modo),
            CuentaDesde = cuentaDesde?.Trim() ?? string.Empty,
            CuentaHasta = cuentaHasta?.Trim() ?? string.Empty,
            OrigenDesde = string.Empty,
            OrigenHasta = string.Empty,
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            PeriodosDisponibles = ListarPeriodosContables(),
            ModosDisponibles =
            [
                new OpcionCatalogoViewModel { Valor = "A", Texto = "Diario auxiliar" },
                new OpcionCatalogoViewModel { Valor = "D", Texto = "Por Cuenta" },
                new OpcionCatalogoViewModel { Valor = "R", Texto = "Por Origen" }
            ]
        };

        if (!consultar)
        {
            return View(model);
        }

        if (EsRangoCuentaInvalido(model.CuentaDesde, model.CuentaHasta))
        {
            model.MensajeError = "Cuenta hasta debe ser mayor o igual que Cuenta desde.";
            return View(model);
        }

        try
        {
            var resultados = await libroDiarioRepository.ListarAsync(new LibroDiarioRequest
            {
                IdEmpresa = idEmpresa,
                Periodo = model.PeriodoConsulta,
                Moneda = "PEN",
                Modo = model.ModoSeleccionado,
                CuentaDesde = model.CuentaDesde,
                CuentaHasta = model.CuentaHasta
            }, cancellationToken);

            model.Resultados = resultados
                .Select(x => new LibroDiarioItemViewModel
                {
                    Modo = x.Modo,
                    CodigoOrigen = x.CodigoOrigen,
                    NombreOrigen = x.NombreOrigen,
                    Periodo = x.Periodo,
                    NumeroAsiento = x.NumeroAsiento,
                    Item = x.Item,
                    FechaEmision = x.FechaEmision,
                    CodigoCuenta = x.CodigoCuenta,
                    NombreCuenta = x.NombreCuenta,
                    NumeroDocumento = x.NumeroDocumento,
                    NombreAuxiliar = x.NombreAuxiliar,
                    TipoDocumento = x.TipoDocumento,
                    Serie = x.Serie,
                    Referencia = x.Referencia,
                    Glosa = x.Glosa,
                    TipoCambio = x.TipoCambio,
                    Debe = x.Debe,
                    Haber = x.Haber,
                    DebeDolares = x.DebeDolares,
                    HaberDolares = x.HaberDolares
                })
                .ToList();

            model.TotalDebe = model.Resultados.Sum(x => x.Debe);
            model.TotalHaber = model.Resultados.Sum(x => x.Haber);
            model.TotalDebeDolares = model.Resultados.Sum(x => x.DebeDolares);
            model.TotalHaberDolares = model.Resultados.Sum(x => x.HaberDolares);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> LibroMayor(
        short? anio = null,
        byte? mes = null,
        string? cuentaDesde = null,
        string? cuentaHasta = null,
        string? numeroDocumento = null,
        bool consultar = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var fechaDesdeTrabajo = new DateOnly(anioTrabajo, mesTrabajo, 1);
        var fechaHastaTrabajo = new DateOnly(anioTrabajo, mesTrabajo, DateTime.DaysInMonth(anioTrabajo, mesTrabajo));

        var model = new LibroMayorViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            PeriodoConsulta = $"{anioTrabajo:0000}{mesTrabajo:00}",
            CuentaDesde = cuentaDesde?.Trim() ?? string.Empty,
            CuentaHasta = cuentaHasta?.Trim() ?? string.Empty,
            NumeroDocumento = numeroDocumento?.Trim() ?? string.Empty,
            ConsultaEjecutada = consultar,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).Select(x => (short)x).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };

        if (!consultar)
        {
            return View(model);
        }

        if (EsRangoCuentaInvalido(model.CuentaDesde, model.CuentaHasta))
        {
            model.MensajeError = "Cuenta hasta debe ser mayor o igual que Cuenta desde.";
            return View(model);
        }

        try
        {
            var resultados = await libroMayorRepository.ListarAsync(new LibroMayorRequest
            {
                IdEmpresa = idEmpresa,
                Periodo = model.PeriodoConsulta,
                CuentaDesde = model.CuentaDesde,
                CuentaHasta = model.CuentaHasta,
                NumeroDocumento = model.NumeroDocumento
            }, cancellationToken);

            var saldosPorCuenta = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            var saldosPorCuentaDolares = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
            model.Resultados = resultados
                .Select(x =>
                {
                    var saldoActual = saldosPorCuenta.TryGetValue(x.CodigoCuenta, out var saldoRegistrado)
                        ? saldoRegistrado
                        : x.SaldoInicial;
                    var saldoActualDolares = saldosPorCuentaDolares.TryGetValue(x.CodigoCuenta, out var saldoRegistradoDolares)
                        ? saldoRegistradoDolares
                        : x.SaldoInicialDolares;

                    if (x.EsSaldoInicial)
                    {
                        saldoActual = x.SaldoInicial;
                        saldoActualDolares = x.SaldoInicialDolares;
                    }
                    else
                    {
                        saldoActual += x.Debe - x.Haber;
                        saldoActualDolares += x.DebeDolares - x.HaberDolares;
                    }

                    saldosPorCuenta[x.CodigoCuenta] = saldoActual;
                    saldosPorCuentaDolares[x.CodigoCuenta] = saldoActualDolares;

                    return new LibroMayorItemViewModel
                    {
                        CodigoCuenta = x.CodigoCuenta,
                        NombreCuenta = x.NombreCuenta,
                        CodigoOrigen = x.CodigoOrigen,
                        NombreOrigen = x.NombreOrigen,
                        Periodo = x.Periodo,
                        NumeroAsiento = x.NumeroAsiento,
                        Item = x.Item,
                        FechaEmision = x.FechaEmision,
                        TipoDocumento = x.TipoDocumento,
                        Serie = x.Serie,
                        Referencia = x.Referencia,
                        NumeroDocumento = x.NumeroDocumento,
                        NombreAuxiliar = x.NombreAuxiliar,
                        Glosa = x.Glosa,
                        TipoCambio = x.TipoCambio,
                        Debe = x.Debe,
                        Haber = x.Haber,
                        DebeDolares = x.DebeDolares,
                        HaberDolares = x.HaberDolares,
                        SaldoInicial = x.SaldoInicial,
                        SaldoInicialDolares = x.SaldoInicialDolares,
                        EsSaldoInicial = x.EsSaldoInicial,
                        SaldoAcumulado = saldoActual,
                        SaldoAcumuladoDolares = saldoActualDolares
                    };
                })
                .ToList();

            model.TotalDebe = model.Resultados.Where(x => !x.EsSaldoInicial).Sum(x => x.Debe);
            model.TotalHaber = model.Resultados.Where(x => !x.EsSaldoInicial).Sum(x => x.Haber);
            model.TotalDebeDolares = model.Resultados.Where(x => !x.EsSaldoInicial).Sum(x => x.DebeDolares);
            model.TotalHaberDolares = model.Resultados.Where(x => !x.EsSaldoInicial).Sum(x => x.HaberDolares);
        }
        catch (SqlException ex)
        {
            model.MensajeError = ex.Message;
        }
        catch (InvalidOperationException ex)
        {
            model.MensajeError = ex.Message;
        }

        return View(model);
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var hoy = DateTime.Today;
        return (anio ?? (short)hoy.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)hoy.Month);
    }

    private static (short anio, byte periodo) NormalizarPeriodoContable(short? anio, byte? periodo)
    {
        var hoy = DateTime.Today;
        return (anio ?? (short)hoy.Year, periodo is <= 14 ? periodo.Value : (byte)hoy.Month);
    }

    private static (short anio, byte periodoDesde, byte periodoHasta) NormalizarRangoPeriodosContables(short? anio, byte? periodoDesde, byte? periodoHasta)
    {
        var hoy = DateTime.Today;
        var anioTrabajo = anio ?? (short)hoy.Year;
        var desdeTrabajo = periodoDesde is <= 14 ? periodoDesde.Value : (byte)0;
        var hastaTrabajo = periodoHasta is <= 14 ? periodoHasta.Value : (byte)hoy.Month;

        if (hastaTrabajo < desdeTrabajo)
        {
            (desdeTrabajo, hastaTrabajo) = (hastaTrabajo, desdeTrabajo);
        }

        return (anioTrabajo, desdeTrabajo, hastaTrabajo);
    }

    private static string NormalizarMoneda(string? moneda)
    {
        return string.Equals(moneda?.Trim(), "USD", StringComparison.OrdinalIgnoreCase) ? "USD" : "PEN";
    }

    private static bool EsRangoCuentaInvalido(string? cuentaDesde, string? cuentaHasta)
    {
        return !string.IsNullOrWhiteSpace(cuentaDesde)
            && !string.IsNullOrWhiteSpace(cuentaHasta)
            && string.Compare(cuentaDesde.Trim(), cuentaHasta.Trim(), StringComparison.OrdinalIgnoreCase) > 0;
    }

    private static string NormalizarEstado(string? estado)
    {
        return estado?.Trim().ToUpperInvariant() switch
        {
            "P" => "P",
            "C" => "C",
            _ => "T"
        };
    }

    private static string NormalizarTipo(string? tipo)
    {
        return tipo is "1" or "2" ? tipo : "0";
    }

    private static string NormalizarModoLibroDiario(string? modo)
    {
        return modo?.Trim().ToUpperInvariant() switch
        {
            "D" => "D",
            "R" => "R",
            _ => "A"
        };
    }

    private static byte NormalizarGrado(byte? grado, int gradoMaximo)
    {
        var gradoTrabajo = grado is >= 1 ? grado.Value : (byte)1;
        return (byte)Math.Clamp(gradoTrabajo, 1, Math.Max(1, gradoMaximo));
    }

    private static List<MesOpcionViewModel> ListarMesesCalendario()
    {
        return Enumerable.Range(1, 12)
            .Select(x => new MesOpcionViewModel
            {
                Valor = (byte)x,
                Nombre = new DateTime(2000, x, 1).ToString("MMMM")
            })
            .ToList();
    }

    private static List<MesOpcionViewModel> ListarPeriodosContables()
    {
        var periodos = new List<MesOpcionViewModel>
        {
            new() { Valor = 0, Nombre = "Apertura" }
        };

        periodos.AddRange(Enumerable.Range(1, 12)
            .Select(x => new MesOpcionViewModel
            {
                Valor = (byte)x,
                Nombre = new DateTime(2000, x, 1).ToString("MMMM")
            }));

        periodos.Add(new MesOpcionViewModel { Valor = 13, Nombre = "Ajustes y Liquidaciones" });
        periodos.Add(new MesOpcionViewModel { Valor = 14, Nombre = "Cierre de Inventario" });

        return periodos;
    }
}
