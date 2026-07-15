using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("ASIENTOS")]
public class AsientoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IAsientoRepository asientoRepository,
    IConfiguracionContabilizacionRepository configuracionContabilizacionRepository,
    IOrigenRepository origenRepository,
    ICentroCostoRepository centroCostoRepository,
    IPlanCuentaRepository planCuentaRepository,
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository,
    IPersonaRepository personaRepository,
    ICompraRepository compraRepository,
    IVentaRepository ventaRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;
    private const int TamanoAyudaPersona = 20;
    private const byte MesContableMinimo = 0;
    private const byte MesContableMaximo = 15;
    private static readonly string[] ModulosOrigenAutomaticoConfigurado = ["COM", "VEN", "ING", "EGR"];

    [HttpGet]
    public async Task<IActionResult> Index(short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var centrosCosto = (await centroCostoRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken))
            .OrderBy(x => x.CodigoCentroCosto)
            .ToList();
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var asientos = await asientoRepository.ListarPaginadoPorEmpresaAsync(empresaId, anioTrabajo, mesTrabajo, textoBusqueda, pagina, TamanoPagina, false, cancellationToken);
        var origenesAutomaticos = await ObtenerOrigenesAutomaticosConfiguradosAsync(empresaId, cancellationToken);
        var asientosNormalizados = AplicarOrigenesAutomaticos(asientos.Items, origenesAutomaticos);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            origenes,
            monedas,
            cuentas,
            centrosCosto,
            tiposDocumento,
            asientosNormalizados,
            null);
        model.TotalAsientos = asientos.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = asientos.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idAsiento, string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, idAsiento, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idAsiento, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);

        try
        {
            var origenesAutomaticos = await ObtenerOrigenesAutomaticosConfiguradosAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var asiento = AplicarOrigenesAutomaticos(await asientoRepository.ObtenerAsync(idAsiento, cancellationToken), origenesAutomaticos);
            if (asiento is not null && !asiento.PermiteRegistroManual)
            {
                throw new InvalidOperationException("El asiento fue generado automaticamente. Debe eliminarlo desde su modulo de origen.");
            }

            await asientoRepository.EliminarAsync(idAsiento, currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            TempData["AsientoOk"] = "Asiento eliminado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["AsientoError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(AsientoFormViewModel formulario, string? periodo = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;
        var periodoTrabajo = NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = DescomponerPeriodo(periodoTrabajo);
        var fechaContabilizacion = ParsePeriodo(periodoTrabajo);
        formulario.FechaAsiento = fechaContabilizacion;

        NormalizarFormulario(formulario);

        if (formulario.IdAsiento.HasValue)
        {
            var origenesAutomaticos = await ObtenerOrigenesAutomaticosConfiguradosAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var asientoExistente = AplicarOrigenesAutomaticos(await asientoRepository.ObtenerAsync(formulario.IdAsiento.Value, cancellationToken), origenesAutomaticos);
            if (asientoExistente is not null && !asientoExistente.PermiteRegistroManual)
            {
                ModelState.AddModelError(string.Empty, "El asiento fue generado automaticamente y solo puede modificarse desde su modulo de origen.");
            }
        }

        var cuentasMovimiento = (await planCuentaRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, true, cancellationToken))
            .Where(x => x.Estado && x.AceptaMovimiento)
            .ToDictionary(x => x.IdPlanCuenta);
        var centrosCostoActivos = (await centroCostoRepository.ListarPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, true, cancellationToken))
            .GroupBy(x => x.CodigoCentroCosto, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

        ValidarFormulario(formulario, cuentasMovimiento, centrosCostoActivos);

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var result = await asientoRepository.GuardarManualAsync(new GuardarAsientoManualRequest
            {
                IdAsiento = formulario.IdAsiento,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdOrigen = formulario.IdOrigen!.Value,
                Periodo = periodoTrabajo,
                FechaEmision = formulario.FechaEmision,
                FechaAsiento = fechaContabilizacion,
                Glosa = formulario.Glosa.Trim(),
                IdMoneda = formulario.IdMoneda!.Value,
                TipoCambio = formulario.TipoCambio,
                ReferenciaExterna = string.IsNullOrWhiteSpace(formulario.ReferenciaExterna) ? null : formulario.ReferenciaExterna.Trim(),
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarAsientoDetalleRequest
                    {
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        Dh = x.Debe > 0 ? "D" : "H",
                        GlosaDetalle = string.IsNullOrWhiteSpace(x.GlosaDetalle) ? null : x.GlosaDetalle.Trim(),
                        CodigoCentroCosto = string.IsNullOrWhiteSpace(x.CodigoCentroCosto) ? null : x.CodigoCentroCosto.Trim(),
                        TipoDocumento = string.IsNullOrWhiteSpace(x.TipoDocumento) ? null : x.TipoDocumento.Trim(),
                        NumeroDocumento = string.IsNullOrWhiteSpace(x.NumeroDocumento) ? null : x.NumeroDocumento.Trim(),
                        Serie = string.IsNullOrWhiteSpace(x.Serie) ? null : x.Serie.Trim(),
                        TipoCambioLinea = x.TipoCambioLinea,
                        Debe = x.Debe,
                        Haber = x.Haber,
                        ReferenciaLinea = string.IsNullOrWhiteSpace(x.ReferenciaLinea) ? null : x.ReferenciaLinea.Trim()
                    })
                    .ToList()
            }, cancellationToken);

            TempData["AsientoOk"] = $"Asiento {result.Periodo}-{result.NumeroAsiento} guardado correctamente.";
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    [HttpGet]
    public async Task<IActionResult> BuscarPersonasAyuda(string? textoBusqueda = null, int numeroPagina = 1, int tamanoPagina = TamanoAyudaPersona, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "No existe una empresa activa en la sesion." });
        }

        var resultado = await personaRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            textoBusqueda,
            null,
            false,
            false,
            numeroPagina <= 0 ? 1 : numeroPagina,
            tamanoPagina <= 0 ? TamanoAyudaPersona : tamanoPagina,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idPersona = x.IdPersona,
                tipoPersona = x.TipoPersona,
                tipoDocumento = x.TipoDocumento,
                nombreTipoDocumento = x.NombreTipoDocumento,
                numeroDocumento = x.NumeroDocumento,
                nombreCompleto = x.NombreCompleto
            }),
            totalRegistros = resultado.TotalRecords
        });
    }

    [HttpGet]
    public async Task<IActionResult> BuscarComprobantesPersonaAyuda(string? numeroDocumento = null, string? textoBusqueda = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "No existe una empresa activa en la sesion." });
        }

        var numeroDocumentoTrabajo = string.IsNullOrWhiteSpace(numeroDocumento)
            ? null
            : numeroDocumento.Trim();

        if (string.IsNullOrWhiteSpace(numeroDocumentoTrabajo))
        {
            return Json(new { ok = false, mensaje = "Seleccione primero una persona o ingrese un RUC/DNI." });
        }

        var filtroTrabajo = string.IsNullOrWhiteSpace(textoBusqueda) ? null : textoBusqueda.Trim();
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var compras = await compraRepository.ListarPorEmpresaAsync(empresaId, null, cancellationToken);
        var ventas = await ventaRepository.ListarPorEmpresaAsync(empresaId, null, cancellationToken);

        var items = compras
            .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
            .Select(x => new ComprobanteSaldoAyudaDto
            {
                ModuloOperacion = "Compra",
                IdRegistro = x.IdCompra,
                FechaEmision = x.FechaEmision,
                NombrePersona = x.NombreProveedor,
                NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                TipoComprobante = x.TipoComprobante,
                DescripcionTipoComprobante = x.DescripcionTipoComprobante,
                Serie = x.Serie,
                Numero = x.Numero,
                CodigoMoneda = x.CodigoMoneda,
                ImporteTotal = x.ImporteTotal,
                Saldo = x.Saldo
            })
            .Concat(ventas
                .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
                .Select(x => new ComprobanteSaldoAyudaDto
                {
                    ModuloOperacion = "Venta",
                    IdRegistro = x.IdVenta,
                    FechaEmision = x.FechaEmision,
                    NombrePersona = x.NombreCliente,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    TipoComprobante = x.TipoComprobante,
                    DescripcionTipoComprobante = x.DescripcionTipoComprobante,
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoMoneda = x.CodigoMoneda,
                    ImporteTotal = x.ImporteTotal,
                    Saldo = x.Saldo
                }))
            .Where(x => filtroTrabajo is null
                || x.NombrePersona.Contains(filtroTrabajo, StringComparison.OrdinalIgnoreCase)
                || x.DescripcionTipoComprobante.Contains(filtroTrabajo, StringComparison.OrdinalIgnoreCase)
                || x.Serie.Contains(filtroTrabajo, StringComparison.OrdinalIgnoreCase)
                || x.Numero.Contains(filtroTrabajo, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(x => x.FechaEmision)
            .ThenByDescending(x => x.IdRegistro)
            .Take(100)
            .ToList();

        return Json(new
        {
            ok = true,
            items = items.Select(x => new
            {
                moduloOperacion = x.ModuloOperacion,
                idRegistro = x.IdRegistro,
                fechaEmision = x.FechaEmision.ToString("dd/MM/yyyy"),
                nombrePersona = x.NombrePersona,
                numeroDocumentoPersona = x.NumeroDocumentoPersona,
                tipoComprobante = x.TipoComprobante,
                descripcionTipoComprobante = x.DescripcionTipoComprobante,
                serie = x.Serie,
                numero = x.Numero,
                codigoMoneda = x.CodigoMoneda,
                importeTotal = x.ImporteTotal,
                saldo = x.Saldo
            })
        });
    }

    private async Task<IActionResult> CargarFormularioAsync(string? periodo, int? idAsiento, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var centrosCosto = (await centroCostoRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken))
            .OrderBy(x => x.CodigoCentroCosto)
            .ToList();
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var origenesAutomaticos = await ObtenerOrigenesAutomaticosConfiguradosAsync(empresaId, cancellationToken);
        var asientos = AplicarOrigenesAutomaticos(
            await asientoRepository.ListarPorEmpresaAsync(empresaId, periodoTrabajo, true, cancellationToken),
            origenesAutomaticos);
        var asientoEditar = idAsiento.HasValue
            ? AplicarOrigenesAutomaticos(await asientoRepository.ObtenerAsync(idAsiento.Value, cancellationToken), origenesAutomaticos)
            : null;

        if (asientoEditar is not null && asientoEditar.IdEmpresa != empresaId)
        {
            asientoEditar = null;
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            short.Parse(periodoTrabajo[..4]),
            byte.Parse(periodoTrabajo[4..]),
            null,
            origenes,
            monedas,
            cuentas,
            centrosCosto,
            tiposDocumento,
            asientos,
            asientoEditar);

        return View("Formulario", model);
    }

    private async Task<AsientoIndexViewModel> ConstruirViewModelErrorAsync(int empresaId, string periodo, AsientoFormViewModel formulario, CancellationToken cancellationToken)
    {
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var centrosCosto = (await centroCostoRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken))
            .OrderBy(x => x.CodigoCentroCosto)
            .ToList();
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var origenesAutomaticos = await ObtenerOrigenesAutomaticosConfiguradosAsync(empresaId, cancellationToken);
        var asientos = AplicarOrigenesAutomaticos(
            await asientoRepository.ListarPorEmpresaAsync(empresaId, periodo, true, cancellationToken),
            origenesAutomaticos);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodo,
            short.Parse(periodo[..4]),
            byte.Parse(periodo[4..]),
            null,
            origenes,
            monedas,
            cuentas,
            centrosCosto,
            tiposDocumento,
            asientos,
            null);

        model.Formulario = formulario;
        HidratarDetallesFormulario(model.Formulario, cuentas, centrosCosto);
        return model;
    }

    private async Task<HashSet<int>> ObtenerOrigenesAutomaticosConfiguradosAsync(int idEmpresa, CancellationToken cancellationToken)
    {
        var configuracion = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(idEmpresa, cancellationToken);

        return configuracion.Provisiones
            .Where(x => x.Activo
                && x.GeneraAsientoAutomatico
                && x.IdOrigen.HasValue
                && ModulosOrigenAutomaticoConfigurado.Contains(x.ModuloOperacion, StringComparer.OrdinalIgnoreCase))
            .Select(x => x.IdOrigen!.Value)
            .ToHashSet();
    }

    private static IReadOnlyCollection<AsientoResumenDto> AplicarOrigenesAutomaticos(
        IReadOnlyCollection<AsientoResumenDto> asientos,
        IReadOnlySet<int> origenesAutomaticos)
    {
        return asientos.Select(x => AplicarOrigenesAutomaticos(x, origenesAutomaticos)).ToList();
    }

    private static AsientoDto? AplicarOrigenesAutomaticos(AsientoDto? asiento, IReadOnlySet<int> origenesAutomaticos)
    {
        if (asiento is null)
        {
            return null;
        }

        return new AsientoDto
        {
            IdAsiento = asiento.IdAsiento,
            IdEmpresa = asiento.IdEmpresa,
            IdOrigen = asiento.IdOrigen,
            CodigoOrigen = asiento.CodigoOrigen,
            NombreOrigen = asiento.NombreOrigen,
            PermiteRegistroManual = !origenesAutomaticos.Contains(asiento.IdOrigen) && asiento.PermiteRegistroManual,
            Ejercicio = asiento.Ejercicio,
            Mes = asiento.Mes,
            Periodo = asiento.Periodo,
            NumeroAsiento = asiento.NumeroAsiento,
            FechaEmision = asiento.FechaEmision,
            FechaAsiento = asiento.FechaAsiento,
            Glosa = asiento.Glosa,
            IdMoneda = asiento.IdMoneda,
            CodigoMoneda = asiento.CodigoMoneda,
            NombreMoneda = asiento.NombreMoneda,
            SimboloMoneda = asiento.SimboloMoneda,
            TipoCambio = asiento.TipoCambio,
            TotalDebe = asiento.TotalDebe,
            TotalHaber = asiento.TotalHaber,
            TotalImporteS = asiento.TotalImporteS,
            TotalImporteD = asiento.TotalImporteD,
            Estado = asiento.Estado,
            ReferenciaExterna = asiento.ReferenciaExterna,
            Observacion = asiento.Observacion,
            Detalles = asiento.Detalles
        };
    }

    private static AsientoResumenDto AplicarOrigenesAutomaticos(AsientoResumenDto asiento, IReadOnlySet<int> origenesAutomaticos)
    {
        return new AsientoResumenDto
        {
            IdAsiento = asiento.IdAsiento,
            IdEmpresa = asiento.IdEmpresa,
            IdOrigen = asiento.IdOrigen,
            CodigoOrigen = asiento.CodigoOrigen,
            NombreOrigen = asiento.NombreOrigen,
            PermiteRegistroManual = !origenesAutomaticos.Contains(asiento.IdOrigen) && asiento.PermiteRegistroManual,
            Ejercicio = asiento.Ejercicio,
            Mes = asiento.Mes,
            Periodo = asiento.Periodo,
            NumeroAsiento = asiento.NumeroAsiento,
            FechaEmision = asiento.FechaEmision,
            FechaAsiento = asiento.FechaAsiento,
            Glosa = asiento.Glosa,
            IdMoneda = asiento.IdMoneda,
            CodigoMoneda = asiento.CodigoMoneda,
            NombreMoneda = asiento.NombreMoneda,
            SimboloMoneda = asiento.SimboloMoneda,
            TipoCambio = asiento.TipoCambio,
            TotalDebe = asiento.TotalDebe,
            TotalHaber = asiento.TotalHaber,
            TotalImporteS = asiento.TotalImporteS,
            TotalImporteD = asiento.TotalImporteD,
            Estado = asiento.Estado,
            ReferenciaExterna = asiento.ReferenciaExterna,
            Observacion = asiento.Observacion
        };
    }

    private static void NormalizarFormulario(AsientoFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue
                     || !string.IsNullOrWhiteSpace(x.GlosaDetalle)
                     || !string.IsNullOrWhiteSpace(x.CodigoCentroCosto)
                     || !string.IsNullOrWhiteSpace(x.TipoDocumento)
                     || !string.IsNullOrWhiteSpace(x.NumeroDocumento)
                     || !string.IsNullOrWhiteSpace(x.Serie)
                     || !string.IsNullOrWhiteSpace(x.ReferenciaLinea)
                     || x.TipoCambioLinea > 0
                     || x.Debe > 0
                     || x.Haber > 0)
            .Select((x, index) =>
            {
                x.Item = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private async Task<List<OpcionCatalogoViewModel>> ObtenerTiposDocumentoAsync(CancellationToken cancellationToken)
    {
        var tipos = await tipoComprobanteRepository.ListarActivosAsync(false, false, cancellationToken);

        return tipos
            .OrderBy(x => x.CodigoTipoComprobante)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoTipoComprobante,
                Texto = $"{x.CodigoTipoComprobante} - {x.Descripcion}"
            })
            .ToList();
    }

    private void ValidarFormulario(
        AsientoFormViewModel formulario,
        IReadOnlyDictionary<int, PlanCuentaDto> cuentasMovimiento,
        IReadOnlyDictionary<string, CentroCostoDto> centrosCostoActivos)
    {
        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos una linea en el asiento.");
            return;
        }

        decimal totalDebe = 0;
        decimal totalHaber = 0;

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (!detalle.IdPlanCuenta.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione una cuenta.");
            }
            else if (!cuentasMovimiento.TryGetValue(detalle.IdPlanCuenta.Value, out var cuenta))
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "La cuenta seleccionada no esta activa o no acepta movimiento.");
            }
            else
            {
                detalle.RequiereCentroCostoCuenta = cuenta.RequiereCentroCosto;

                if (cuenta.RequiereCentroCosto && string.IsNullOrWhiteSpace(detalle.CodigoCentroCosto))
                {
                    ModelState.AddModelError($"{prefijo}.CodigoCentroCosto", "La cuenta seleccionada requiere centro de costo.");
                }
            }

            if (!string.IsNullOrWhiteSpace(detalle.CodigoCentroCosto)
                && !centrosCostoActivos.ContainsKey(detalle.CodigoCentroCosto.Trim()))
            {
                ModelState.AddModelError($"{prefijo}.CodigoCentroCosto", "El centro de costo ingresado no existe o no esta activo para la empresa.");
            }

            if (detalle.TipoCambioLinea <= 0)
            {
                ModelState.AddModelError($"{prefijo}.TipoCambioLinea", "Ingrese un tipo de cambio mayor a cero en la linea.");
            }

            var tieneDebe = detalle.Debe > 0;
            var tieneHaber = detalle.Haber > 0;

            if (tieneDebe == tieneHaber)
            {
                ModelState.AddModelError($"{prefijo}.Debe", "La linea debe tener monto solo en Debe o solo en Haber.");
            }

            totalDebe += detalle.Debe;
            totalHaber += detalle.Haber;
        }

        if (totalDebe <= 0 && totalHaber <= 0)
        {
            ModelState.AddModelError(string.Empty, "El asiento debe tener al menos un importe positivo en el detalle.");
        }
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var today = DateTime.Today;
        var anioTrabajo = anio ?? (short)today.Year;
        var mesTrabajo = EsMesContableValido(mes) ? mes!.Value : (byte)today.Month;
        return (anioTrabajo, mesTrabajo);
    }

    private static string NormalizarPeriodo(string? periodo)
    {
        if (!string.IsNullOrWhiteSpace(periodo)
            && periodo.Length == 6
            && short.TryParse(periodo[..4], out var anio)
            && byte.TryParse(periodo[4..], out var mes)
            && EsMesContableValido(mes))
        {
            return $"{anio:0000}{mes:00}";
        }

        var (anioActual, mesActual) = NormalizarPeriodo(null, null);
        return $"{anioActual:0000}{mesActual:00}";
    }

    private static AsientoIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        string periodo,
        short anioSeleccionado,
        byte mesSeleccionado,
        string? textoBusqueda,
        IReadOnlyCollection<OrigenDto> origenes,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        IReadOnlyCollection<CentroCostoDto> centrosCosto,
        IReadOnlyCollection<OpcionCatalogoViewModel> tiposDocumento,
        IReadOnlyCollection<AsientoResumenDto> asientos,
        AsientoDto? asientoEditar)
    {
        var items = asientos
            .Select(x => new AsientoResumenItemViewModel
            {
                IdAsiento = x.IdAsiento,
                CodigoOrigen = x.CodigoOrigen,
                NombreOrigen = x.NombreOrigen,
                Periodo = x.Periodo,
                NumeroAsiento = x.NumeroAsiento,
                FechaEmision = x.FechaEmision,
                FechaAsiento = x.FechaAsiento,
                Glosa = x.Glosa,
                CodigoMoneda = x.CodigoMoneda,
                TipoCambio = x.TipoCambio,
                TotalDebe = x.TotalDebe,
                TotalHaber = x.TotalHaber,
                TotalImporteS = x.TotalImporteS,
                TotalImporteD = x.TotalImporteD,
                Estado = x.Estado,
                PermiteRegistroManual = x.PermiteRegistroManual
            })
            .ToList();
        var totalDebePeriodo = items.Sum(x => x.TotalDebe);
        var totalDebeSolesPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "PEN", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.TotalDebe);
        var totalDebeDolaresPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "USD", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.TotalDebe);
        var totalHaberPeriodo = items.Sum(x => x.TotalHaber);
        var totalHaberSolesPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "PEN", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.TotalHaber);
        var totalHaberDolaresPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "USD", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.TotalHaber);

        return new AsientoIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TotalAsientos = items.Count,
            TotalDebePeriodo = totalDebePeriodo,
            TotalDebeSolesPeriodo = totalDebeSolesPeriodo,
            TotalDebeDolaresPeriodo = totalDebeDolaresPeriodo,
            TotalHaberPeriodo = totalHaberPeriodo,
            TotalHaberSolesPeriodo = totalHaberSolesPeriodo,
            TotalHaberDolaresPeriodo = totalHaberDolaresPeriodo,
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            OrigenesManual = origenes.ToList(),
            Monedas = monedas.ToList(),
            CuentasMovimiento = cuentas.ToList(),
            TiposDocumentoIdentidad = tiposDocumento.ToList(),
            Asientos = items,
            Formulario = asientoEditar is null
                ? new AsientoFormViewModel
                {
                    OrigenTexto = origenes.FirstOrDefault() is { } origenDefault
                        ? $"{origenDefault.CodigoOrigen} - {origenDefault.NombreOrigen}"
                        : string.Empty,
                    FechaEmision = ParsePeriodo(periodo),
                    FechaAsiento = ParsePeriodo(periodo),
                    IdOrigen = origenes.FirstOrDefault()?.IdOrigen,
                    IdMoneda = monedas.OrderByDescending(x => x.EsMonedaBase).FirstOrDefault()?.IdMoneda
                }
                : new AsientoFormViewModel
                {
                    IdAsiento = asientoEditar.IdAsiento,
                    NumeroAsiento = asientoEditar.NumeroAsiento,
                    PermiteRegistroManual = asientoEditar.PermiteRegistroManual,
                    IdOrigen = asientoEditar.IdOrigen,
                    OrigenTexto = $"{asientoEditar.CodigoOrigen} - {asientoEditar.NombreOrigen}",
                    FechaEmision = asientoEditar.FechaEmision,
                    FechaAsiento = asientoEditar.FechaAsiento,
                    Glosa = asientoEditar.Glosa,
                    IdMoneda = asientoEditar.IdMoneda,
                    TipoCambio = asientoEditar.TipoCambio,
                    ReferenciaExterna = asientoEditar.ReferenciaExterna,
                    Observacion = asientoEditar.Observacion,
                    Detalles = asientoEditar.Detalles
                        .OrderBy(x => x.Item)
                        .Select(x => new AsientoDetalleFormViewModel
                        {
                            IdAsientoDetalle = x.IdAsientoDetalle,
                            Item = x.Item,
                            IdPlanCuenta = x.IdPlanCuenta,
                            CuentaTexto = $"{x.CodigoCuenta} - {x.NombreCuenta}",
                            RequiereCentroCostoCuenta = cuentas.FirstOrDefault(c => c.IdPlanCuenta == x.IdPlanCuenta)?.RequiereCentroCosto ?? false,
                            GlosaDetalle = x.GlosaDetalle,
                            CodigoCentroCosto = x.CodigoCentroCosto,
                            CentroCostoTexto = ObtenerCentroCostoTexto(centrosCosto, x.CodigoCentroCosto),
                            TipoDocumento = x.TipoDocumento,
                            NumeroDocumento = x.NumeroDocumento,
                            PersonaTexto = x.NumeroDocumento ?? string.Empty,
                            Serie = x.Serie,
                            TipoCambioLinea = x.TipoCambioLinea ?? asientoEditar.TipoCambio,
                            Debe = x.Debe,
                            Haber = x.Haber,
                            ReferenciaLinea = x.ReferenciaLinea
                        })
                        .ToList()
                }
        };
    }

    private static void HidratarDetallesFormulario(
        AsientoFormViewModel formulario,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        IReadOnlyCollection<CentroCostoDto> centrosCosto)
    {
        var cuentasPorId = cuentas.ToDictionary(x => x.IdPlanCuenta);
        var centrosPorCodigo = centrosCosto
            .GroupBy(x => x.CodigoCentroCosto, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

        foreach (var detalle in formulario.Detalles)
        {
            if (detalle.IdPlanCuenta.HasValue && cuentasPorId.TryGetValue(detalle.IdPlanCuenta.Value, out var cuenta))
            {
                detalle.CuentaTexto = $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}";
                detalle.RequiereCentroCostoCuenta = cuenta.RequiereCentroCosto;
            }

            if (!string.IsNullOrWhiteSpace(detalle.CodigoCentroCosto))
            {
                detalle.CentroCostoTexto = centrosPorCodigo.TryGetValue(detalle.CodigoCentroCosto.Trim(), out var centro)
                    ? $"{centro.CodigoCentroCosto} - {centro.NombreCentroCosto}"
                    : detalle.CodigoCentroCosto.Trim();
            }

            detalle.PersonaTexto = detalle.NumeroDocumento?.Trim() ?? string.Empty;
        }
    }

    private static string ObtenerCentroCostoTexto(IReadOnlyCollection<CentroCostoDto> centrosCosto, string? codigoCentroCosto)
    {
        if (string.IsNullOrWhiteSpace(codigoCentroCosto))
        {
            return string.Empty;
        }

        var centro = centrosCosto.FirstOrDefault(x => string.Equals(x.CodigoCentroCosto, codigoCentroCosto.Trim(), StringComparison.OrdinalIgnoreCase));
        return centro is null
            ? codigoCentroCosto.Trim()
            : $"{centro.CodigoCentroCosto} - {centro.NombreCentroCosto}";
    }

    private static List<int> ConstruirAnios(short anioSeleccionado)
    {
        return Enumerable.Range(anioSeleccionado - 5, 11).ToList();
    }

    private static List<MesOpcionViewModel> ConstruirMeses()
    {
        return
        [
            new() { Valor = 0, Nombre = "Apertura" },
            new() { Valor = 1, Nombre = "Enero" },
            new() { Valor = 2, Nombre = "Febrero" },
            new() { Valor = 3, Nombre = "Marzo" },
            new() { Valor = 4, Nombre = "Abril" },
            new() { Valor = 5, Nombre = "Mayo" },
            new() { Valor = 6, Nombre = "Junio" },
            new() { Valor = 7, Nombre = "Julio" },
            new() { Valor = 8, Nombre = "Agosto" },
            new() { Valor = 9, Nombre = "Setiembre" },
            new() { Valor = 10, Nombre = "Octubre" },
            new() { Valor = 11, Nombre = "Noviembre" },
            new() { Valor = 12, Nombre = "Diciembre" },
            new() { Valor = 13, Nombre = "Ajustes y Liquidaciones" },
            new() { Valor = 14, Nombre = "Cierre de Ganancias y Perdidas" },
            new() { Valor = 15, Nombre = "Cierre de Inventarios" }
        ];
    }

    private static DateOnly ParsePeriodo(string periodo)
    {
        if (periodo.Length == 6
            && int.TryParse(periodo[..4], out var year)
            && int.TryParse(periodo[4..], out var month)
            && month >= MesContableMinimo
            && month <= MesContableMaximo)
        {
            return month switch
            {
                0 => new DateOnly(year, 1, 1),
                >= 1 and <= 12 => new DateOnly(year, month, 1),
                _ => new DateOnly(year, 12, 31)
            };
        }

        return DateOnly.FromDateTime(DateTime.Today);
    }

    private static bool EsMesContableValido(byte? mes)
        => mes.HasValue
           && mes.Value >= MesContableMinimo
           && mes.Value <= MesContableMaximo;

    private static (short anio, byte mes) DescomponerPeriodo(string periodo)
    {
        return (short.Parse(periodo[..4]), byte.Parse(periodo[4..]));
    }
}
