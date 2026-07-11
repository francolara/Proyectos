using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CAJABANCOS")]
public class CajaBancoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPeriodoContableService periodoContableService,
    ICajaBancoRepository cajaBancoRepository,
    IConfiguracionContabilizacionRepository configuracionContabilizacionRepository,
    ICuentaCorrienteRepository cuentaCorrienteRepository,
    IPlanCuentaRepository planCuentaRepository,
    ICentroCostoRepository centroCostoRepository,
    IPersonaRepository personaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository,
    ICompraRepository compraRepository,
    IVentaRepository ventaRepository,
    IParametroEmpresaRepository parametroEmpresaRepository) : Controller
{
    private const int TamanoPagina = 20;

    [HttpGet]
    public async Task<IActionResult> Index(int? idBancoConfiguracionEmpresa = null, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var movimientos = idBancoConfiguracionEmpresa.HasValue
            ? await cajaBancoRepository.ListarPaginadoPorEmpresaAsync(empresaId, idBancoConfiguracionEmpresa, anioTrabajo, mesTrabajo, textoBusqueda, pagina, TamanoPagina, cancellationToken)
            : new PagedResult<CajaBancoResumenDto> { Items = [], TotalRecords = 0, PageNumber = pagina, PageSize = TamanoPagina };
        var resumen = idBancoConfiguracionEmpresa.HasValue
            ? await cajaBancoRepository.ObtenerResumenCuentaAsync(empresaId, idBancoConfiguracionEmpresa, anioTrabajo, mesTrabajo, cancellationToken)
            : new CajaBancoResumenCuentaDto();

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            idBancoConfiguracionEmpresa,
            textoBusqueda,
            cuentas,
            [],
            [],
            tiposDocumento,
            resumen,
            movimientos.Items,
            null);

        model.TotalMovimientos = movimientos.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = movimientos.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(int? idBancoConfiguracionEmpresa = null, short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idBancoConfiguracionEmpresa, anio, mes, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idMovimientoBanco, int? idBancoConfiguracionEmpresa = null, short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idBancoConfiguracionEmpresa, anio, mes, idMovimientoBanco, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> BuscarPersonasAyuda(string? textoBusqueda = null, int numeroPagina = 1, int tamanoPagina = 20, CancellationToken cancellationToken = default)
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
            tamanoPagina <= 0 ? 20 : tamanoPagina,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idPersona = x.IdPersona,
                numeroDocumento = x.NumeroDocumento,
                nombreCompleto = x.NombreCompleto,
                tipoPersona = x.TipoPersona,
                tipoDocumento = x.TipoDocumento,
                nombreTipoDocumento = x.NombreTipoDocumento
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
        var detracciones = await compraRepository.ListarDetraccionesPendientesPorEmpresaAsync(empresaId, null, cancellationToken);
        var percepciones = await compraRepository.ListarPercepcionesPendientesPorEmpresaAsync(empresaId, null, cancellationToken);
        var retenciones = await compraRepository.ListarRetencionesPendientesPorEmpresaAsync(empresaId, null, cancellationToken);
        var ventas = await ventaRepository.ListarPorEmpresaAsync(empresaId, null, cancellationToken);
        var configuracionContable = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(empresaId, cancellationToken);
        var tiposComprobante = await tipoComprobanteRepository.ListarActivosAsync(false, false, cancellationToken);
        var cuentaDetraccion = await ObtenerCuentaDetraccionAsync(empresaId, cancellationToken);
        var cuentaPercepcion = await ObtenerCuentaPercepcionAsync(empresaId, cancellationToken);
        var cuentaRetencion = ObtenerCuentaRetencion(configuracionContable);
        var tipoComprobantePorCodigo = tiposComprobante
            .GroupBy(x => x.CodigoTipoComprobante, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
        var configuracionDocumentoPorId = configuracionContable.Documentos
            .Where(x => x.Activo)
            .ToDictionary(x => x.IdTipoComprobante);

        var items = compras
            .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
            .Select(x => new ComprobanteSaldoAyudaDto
            {
                ModuloOperacion = x.ModuloOperacion,
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
            .Concat(detracciones
                .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
                .Select(x => new ComprobanteSaldoAyudaDto
                {
                    ModuloOperacion = "DET",
                    IdRegistro = x.IdCompraDetraccion,
                    FechaEmision = x.FechaEmision,
                    NombrePersona = x.NombreProveedor,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    TipoComprobante = "00",
                    DescripcionTipoComprobante = "Otros",
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoMoneda = x.CodigoMoneda,
                    ImporteTotal = x.ImporteDetraccion,
                    Saldo = x.Saldo
                }))
            .Concat(percepciones
                .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
                .Select(x => new ComprobanteSaldoAyudaDto
                {
                    ModuloOperacion = "PER",
                    IdRegistro = x.IdCompraPercepcion,
                    FechaEmision = x.FechaEmision,
                    NombrePersona = x.NombreProveedor,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    TipoComprobante = "00",
                    DescripcionTipoComprobante = "Otros",
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoMoneda = x.CodigoMoneda,
                    ImporteTotal = x.ImportePercepcion,
                    Saldo = x.Saldo
                }))
            .Concat(retenciones
                .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
                .Select(x => new ComprobanteSaldoAyudaDto
                {
                    ModuloOperacion = "R4T",
                    IdRegistro = x.IdCompraRetencion,
                    FechaEmision = x.FechaEmision,
                    NombrePersona = x.NombreProveedor,
                    NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                    TipoComprobante = "00",
                    DescripcionTipoComprobante = "Otros",
                    Serie = x.Serie,
                    Numero = x.Numero,
                    CodigoMoneda = x.CodigoMoneda,
                    ImporteTotal = x.Retencion,
                    Saldo = x.Saldo
                }))
            .Concat(ventas
                .Where(x => string.Equals(x.NumeroDocumentoPersona, numeroDocumentoTrabajo, StringComparison.OrdinalIgnoreCase) && x.Saldo > 0)
                .Select(x => new ComprobanteSaldoAyudaDto
                {
                    ModuloOperacion = x.ModuloOperacion,
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
                cuentaContable = ResolverCuentaContableComprobante(x, tipoComprobantePorCodigo, configuracionDocumentoPorId, cuentaDetraccion, cuentaPercepcion, cuentaRetencion),
                moduloOperacion = x.ModuloOperacion,
                moduloOperacionTexto = string.Equals(x.ModuloOperacion, "COM", StringComparison.OrdinalIgnoreCase)
                    ? "Compra"
                    : string.Equals(x.ModuloOperacion, "DET", StringComparison.OrdinalIgnoreCase)
                        ? "Detraccion"
                        : string.Equals(x.ModuloOperacion, "PER", StringComparison.OrdinalIgnoreCase)
                            ? "Percepciones"
                            : string.Equals(x.ModuloOperacion, "R4T", StringComparison.OrdinalIgnoreCase)
                                ? "Renta4ta"
                                : "Venta",
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
            }),
            totalRegistros = items.Count
        });
    }

    [HttpGet]
    public async Task<IActionResult> BuscarComprobantesDetalleAyuda(int? idPersona = null, string? textoBusqueda = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "No existe una empresa activa en la sesion." });
        }

        if (!idPersona.HasValue || idPersona <= 0)
        {
            return Json(new { ok = false, mensaje = "Seleccione una persona en la linea antes de buscar comprobantes." });
        }

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var persona = await personaRepository.ObtenerPorIdAsync(empresaId, idPersona.Value, cancellationToken);
        if (persona is null)
        {
            return Json(new { ok = false, mensaje = "La persona seleccionada no existe o no pertenece a la empresa activa." });
        }

        return await BuscarComprobantesPersonaAyuda(persona.NumeroDocumento, textoBusqueda, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(CajaBancoFormViewModel formulario, short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
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
                (short)formulario.FechaEmision.Year,
                (byte)formulario.FechaEmision.Month,
                cancellationToken))
        {
            ModelState.AddModelError(
                string.Empty,
                periodoContableService.ConstruirMensajeBloqueo(
                    (short)formulario.FechaEmision.Year,
                    (byte)formulario.FechaEmision.Month));
        }

        var cuentasCorrientesActivas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var cuentasCorrientesValidas = cuentasCorrientesActivas
            .ToDictionary(x => x.IdBancoConfiguracionEmpresa);
        var operacionesValidas = (await cajaBancoRepository.ListarOperacionesBancariasAsync(
                string.Equals(formulario.TipoMovimiento, "E", StringComparison.OrdinalIgnoreCase) ? "E" : "I",
                null,
                100,
                null,
                cancellationToken))
            .Select(x => x.IdOpeBancaria.Trim().ToUpperInvariant())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var cuentasMovimiento = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado && x.AceptaMovimiento)
            .ToDictionary(x => x.IdPlanCuenta);
        var centrosCostoActivos = (await centroCostoRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .GroupBy(x => x.CodigoCentroCosto, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);

        var personasDetalle = new Dictionary<int, PersonaDetalleDto>();
        foreach (var idPersonaDetalle in formulario.Detalles
                     .Where(x => x.IdPersona.HasValue && x.IdPersona.Value > 0)
                     .Select(x => x.IdPersona!.Value)
                     .Distinct())
        {
            var personaDetalle = await personaRepository.ObtenerPorIdAsync(empresaId, idPersonaDetalle, cancellationToken);
            if (personaDetalle is not null)
            {
                personasDetalle[idPersonaDetalle] = personaDetalle;
            }
        }

        ValidarFormulario(formulario, cuentasCorrientesValidas, operacionesValidas, cuentasMovimiento, centrosCostoActivos, personasDetalle);

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(empresaId, formulario, anioTrabajo, mesTrabajo, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var result = await cajaBancoRepository.GuardarAsync(new GuardarCajaBancoRequest
            {
                IdMovimientoBanco = formulario.IdMovimientoBanco,
                IdEmpresa = empresaId,
                IdBancoConfiguracionEmpresa = formulario.IdBancoConfiguracionEmpresa!.Value,
                TipoMovimiento = formulario.TipoMovimiento.Trim().ToUpperInvariant(),
                IdOpeBancaria = formulario.IdOpeBancaria.Trim().ToUpperInvariant(),
                FechaEmision = formulario.FechaEmision,
                TipoCambio = formulario.TipoCambio,
                IdPersona = formulario.IdPersona,
                NumeroDocumento = string.IsNullOrWhiteSpace(formulario.NumeroDocumento) ? null : formulario.NumeroDocumento.Trim(),
                Glosa = formulario.Glosa.Trim(),
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                ImporteTotal = formulario.ImporteTotal,
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarCajaBancoDetalleRequest
                    {
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        IdPersona = x.IdPersona,
                        ModuloOperacionComprobante = string.IsNullOrWhiteSpace(x.ModuloOperacionComprobante) ? null : x.ModuloOperacionComprobante.Trim().ToUpperInvariant(),
                        IdRegistroComprobante = x.IdRegistroComprobante,
                        ImporteAplicado = x.Debe > 0 ? x.Debe : x.Haber,
                        GlosaDetalle = string.IsNullOrWhiteSpace(x.GlosaDetalle) ? null : x.GlosaDetalle.Trim(),
                        CodigoCentroCosto = string.IsNullOrWhiteSpace(x.CodigoCentroCosto) ? null : x.CodigoCentroCosto.Trim(),
                        NumeroDocumento = string.IsNullOrWhiteSpace(x.NumeroDocumento) ? null : x.NumeroDocumento.Trim(),
                        TipoDocumento = string.IsNullOrWhiteSpace(x.TipoDocumento) ? null : x.TipoDocumento.Trim(),
                        Serie = string.IsNullOrWhiteSpace(x.Serie) ? null : x.Serie.Trim(),
                        ReferenciaLinea = string.IsNullOrWhiteSpace(x.ReferenciaLinea) ? null : x.ReferenciaLinea.Trim(),
                        TipoCambioLinea = x.TipoCambioLinea,
                        Debe = x.Debe,
                        Haber = x.Haber
                    })
                    .ToList()
            }, cancellationToken);

            TempData["CajaBancoOk"] = result.NumeroAsiento.HasValue
                ? (formulario.IdMovimientoBanco.HasValue
                    ? $"Movimiento de caja y bancos actualizado correctamente. Asiento vinculado: {result.NumeroAsiento.Value}."
                    : $"Movimiento de caja y bancos registrado correctamente. Asiento vinculado: {result.NumeroAsiento.Value}.")
                : (formulario.IdMovimientoBanco.HasValue
                    ? "Movimiento de caja y bancos actualizado correctamente."
                    : "Movimiento de caja y bancos registrado correctamente.");

            var anioResultado = (short)formulario.FechaEmision.Year;
            var mesResultado = (byte)formulario.FechaEmision.Month;

            return RedirectToAction(nameof(Index), new
            {
                idBancoConfiguracionEmpresa = formulario.IdBancoConfiguracionEmpresa,
                anio = anioResultado,
                mes = mesResultado
            });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirViewModelErrorAsync(empresaId, formulario, anioTrabajo, mesTrabajo, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idMovimientoBanco, int? idBancoConfiguracionEmpresa = null, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CajaBancoError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { idBancoConfiguracionEmpresa, anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
        }

        try
        {
            await cajaBancoRepository.EliminarAsync(currentCompanyAccessor.EmpresaId.Value, idMovimientoBanco, cancellationToken);
            TempData["CajaBancoOk"] = "Movimiento de caja y bancos eliminado correctamente.";
        }
        catch (Exception ex)
        {
            TempData["CajaBancoError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { idBancoConfiguracionEmpresa, anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idBancoConfiguracionEmpresa, short? anio, byte? mes, int? idMovimientoBanco, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        if (await periodoContableService.EstaCerradoAsync(empresaId, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CajaBancoError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { idBancoConfiguracionEmpresa, anio = anioTrabajo, mes = mesTrabajo });
        }

        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var resumen = await cajaBancoRepository.ObtenerResumenCuentaAsync(empresaId, idBancoConfiguracionEmpresa, anioTrabajo, mesTrabajo, cancellationToken);
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var operacionesIngreso = (await cajaBancoRepository.ListarOperacionesBancariasAsync("I", null, 100, null, cancellationToken))
            .OrderBy(x => x.IdOpeBancaria)
            .ThenBy(x => x.TipoOperacion)
            .ToList();
        var operacionesEgreso = (await cajaBancoRepository.ListarOperacionesBancariasAsync("E", null, 100, null, cancellationToken))
            .OrderBy(x => x.IdOpeBancaria)
            .ThenBy(x => x.TipoOperacion)
            .ToList();
        var movimientos = await cajaBancoRepository.ListarPaginadoPorEmpresaAsync(empresaId, idBancoConfiguracionEmpresa, anioTrabajo, mesTrabajo, null, 1, TamanoPagina, cancellationToken);
        var movimientoEditar = idMovimientoBanco.HasValue
            ? await cajaBancoRepository.ObtenerPorIdAsync(empresaId, idMovimientoBanco.Value, cancellationToken)
            : null;
        var cuentasPlan = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();

        if (movimientoEditar is not null)
        {
            idBancoConfiguracionEmpresa ??= movimientoEditar.IdBancoConfiguracionEmpresa;
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            idBancoConfiguracionEmpresa,
            null,
            cuentas,
            operacionesIngreso,
            operacionesEgreso,
            tiposDocumento,
            resumen,
            movimientos.Items,
            movimientoEditar);

        HidratarFormulario(model, cuentas, operacionesIngreso, operacionesEgreso, cuentasPlan);
        return View("Formulario", model);
    }

    private async Task<CajaBancoIndexViewModel> ConstruirViewModelErrorAsync(int empresaId, CajaBancoFormViewModel formulario, short anio, byte mes, CancellationToken cancellationToken)
    {
        var periodoTrabajo = $"{anio:0000}{mes:00}";
        var cuentas = (await cuentaCorrienteRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken)).ToList();
        var resumen = await cajaBancoRepository.ObtenerResumenCuentaAsync(empresaId, formulario.IdBancoConfiguracionEmpresa, anio, mes, cancellationToken);
        var tiposDocumento = await ObtenerTiposDocumentoAsync(cancellationToken);
        var operacionesIngreso = (await cajaBancoRepository.ListarOperacionesBancariasAsync("I", null, 100, null, cancellationToken))
            .OrderBy(x => x.IdOpeBancaria)
            .ThenBy(x => x.TipoOperacion)
            .ToList();
        var operacionesEgreso = (await cajaBancoRepository.ListarOperacionesBancariasAsync("E", null, 100, null, cancellationToken))
            .OrderBy(x => x.IdOpeBancaria)
            .ThenBy(x => x.TipoOperacion)
            .ToList();
        var movimientos = await cajaBancoRepository.ListarPaginadoPorEmpresaAsync(empresaId, formulario.IdBancoConfiguracionEmpresa, anio, mes, null, 1, TamanoPagina, cancellationToken);
        var cuentasPlan = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anio,
            mes,
            formulario.IdBancoConfiguracionEmpresa,
            null,
            cuentas,
            operacionesIngreso,
            operacionesEgreso,
            tiposDocumento,
            resumen,
            movimientos.Items,
            null);

        model.Formulario = formulario;
        HidratarFormulario(model, cuentas, operacionesIngreso, operacionesEgreso, cuentasPlan);
        return model;
    }

    private static CajaBancoIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        string periodo,
        short anioSeleccionado,
        byte mesSeleccionado,
        int? idBancoConfiguracionEmpresaSeleccionada,
        string? textoBusqueda,
        IReadOnlyCollection<BancoConfiguracionEmpresaDto> cuentasCorrientes,
        IReadOnlyCollection<OperacionBancariaDto> operacionesIngreso,
        IReadOnlyCollection<OperacionBancariaDto> operacionesEgreso,
        IReadOnlyCollection<OpcionCatalogoViewModel> tiposDocumento,
        CajaBancoResumenCuentaDto resumen,
        IReadOnlyCollection<CajaBancoResumenDto> movimientos,
        CajaBancoMovimientoDto? movimientoEditar)
    {
        var items = movimientos
            .Select(x => new CajaBancoResumenItemViewModel
            {
                IdMovimientoBanco = x.IdMovimientoBanco,
                IdAsiento = x.IdAsiento,
                NumeroAsiento = x.NumeroAsiento,
                NumeroMovimiento = x.NumeroMovimiento,
                FechaEmision = x.FechaEmision,
                IdBancoConfiguracionEmpresa = x.IdBancoConfiguracionEmpresa,
                NroCuentaCorriente = x.NroCuentaCorriente,
                CodigoMoneda = x.CodigoMoneda,
                TipoMovimiento = x.TipoMovimiento,
                TipoOperacion = x.TipoOperacion,
                NombrePersona = x.NombrePersona,
                NumeroOperacion = x.NumeroDocumento,
                Glosa = x.Glosa,
                Ingreso = x.Ingreso,
                Egreso = x.Egreso
            })
            .ToList();

        var formulario = movimientoEditar is null
            ? new CajaBancoFormViewModel
            {
                IdBancoConfiguracionEmpresa = idBancoConfiguracionEmpresaSeleccionada ?? cuentasCorrientes.FirstOrDefault()?.IdBancoConfiguracionEmpresa,
                FechaEmision = ParsePeriodo(periodo),
                TipoCambio = 0m,
                TipoMovimiento = "I",
                IdOpeBancaria = operacionesIngreso.FirstOrDefault()?.IdOpeBancaria ?? string.Empty,
                TipoOperacionTexto = operacionesIngreso.FirstOrDefault()?.TipoOperacion ?? string.Empty,
                Detalles =
                [
                    new CajaBancoDetalleFormViewModel
                    {
                        Item = 1
                    }
                ]
            }
            : new CajaBancoFormViewModel
            {
                IdMovimientoBanco = movimientoEditar.IdMovimientoBanco,
                NumeroMovimiento = movimientoEditar.NumeroMovimiento,
                IdBancoConfiguracionEmpresa = movimientoEditar.IdBancoConfiguracionEmpresa,
                TipoMovimiento = movimientoEditar.TipoMovimiento,
                IdOpeBancaria = movimientoEditar.IdOpeBancaria,
                FechaEmision = movimientoEditar.FechaEmision,
                TipoCambio = movimientoEditar.TipoCambio,
                IdPersona = movimientoEditar.IdPersona,
                PersonaTexto = movimientoEditar.NombrePersona,
                NumeroDocumentoPersona = movimientoEditar.NumeroDocumentoPersona,
                NumeroDocumento = movimientoEditar.NumeroDocumento,
                Glosa = movimientoEditar.Glosa,
                Observacion = movimientoEditar.Observacion,
                ImporteTotal = movimientoEditar.ImporteTotal,
                Detalles = movimientoEditar.Detalles
                    .OrderBy(x => x.Item)
                    .Select(x => new CajaBancoDetalleFormViewModel
                    {
                        IdMovimientoBancoDetalle = x.IdMovimientoBancoDetalle,
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta,
                        IdPersona = x.IdPersona,
                        ModuloOperacionComprobante = x.ModuloOperacionComprobante,
                        IdRegistroComprobante = x.IdRegistroComprobante,
                        PersonaTexto = x.NombrePersona,
                        CuentaTexto = $"{x.CodigoCuenta} - {x.NombreCuenta}",
                        RequiereCentroCostoCuenta = x.RequiereCentroCosto,
                        NumeroDocumento = x.NumeroDocumentoPersona,
                        GlosaDetalle = x.GlosaDetalle,
                        CodigoCentroCosto = x.CodigoCentroCosto,
                        TipoDocumento = x.TipoDocumento,
                        Serie = x.Serie,
                        ReferenciaLinea = x.ReferenciaLinea,
                        TipoCambioLinea = x.TipoCambioLinea ?? movimientoEditar.TipoCambio,
                        Debe = x.Debe,
                        Haber = x.Haber
                    })
                    .ToList()
            };

        return new CajaBancoIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            IdBancoConfiguracionEmpresaSeleccionada = idBancoConfiguracionEmpresaSeleccionada,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            SaldoInicial = resumen.SaldoInicial,
            IngresosMes = resumen.IngresosMes,
            EgresosMes = resumen.EgresosMes,
            SaldoFinal = resumen.SaldoFinal,
            TotalMovimientos = items.Count,
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            CuentasCorrientesDisponibles = cuentasCorrientes.ToList(),
            OperacionesIngreso = operacionesIngreso.Select(x => new CajaBancoOperacionViewModel
            {
                IdOpeBancaria = x.IdOpeBancaria,
                TipoOperacion = x.TipoOperacion
            }).ToList(),
            OperacionesEgreso = operacionesEgreso.Select(x => new CajaBancoOperacionViewModel
            {
                IdOpeBancaria = x.IdOpeBancaria,
                TipoOperacion = x.TipoOperacion
            }).ToList(),
            TiposDocumentoDisponibles = tiposDocumento.ToList(),
            Movimientos = items,
            Formulario = formulario
        };
    }

    private static void HidratarFormulario(
        CajaBancoIndexViewModel model,
        IReadOnlyCollection<BancoConfiguracionEmpresaDto> cuentas,
        IReadOnlyCollection<OperacionBancariaDto> operacionesIngreso,
        IReadOnlyCollection<OperacionBancariaDto> operacionesEgreso,
        IReadOnlyCollection<PlanCuentaDto> cuentasPlan)
    {
        if (model.Formulario.IdBancoConfiguracionEmpresa.HasValue)
        {
            var cuenta = cuentas.FirstOrDefault(x => x.IdBancoConfiguracionEmpresa == model.Formulario.IdBancoConfiguracionEmpresa.Value);
            if (cuenta is not null)
            {
                model.Formulario.CuentaCorrienteTexto = cuenta.NroCuentaCorriente;
                model.Formulario.BancoTexto = $"{cuenta.NombreBanco} ({cuenta.CodigoBanco})";
                model.Formulario.MonedaTexto = string.IsNullOrWhiteSpace(cuenta.CodigoMoneda)
                    ? string.Empty
                    : $"{cuenta.CodigoMoneda} - {cuenta.NombreMoneda}";
            }
        }

        var operaciones = string.Equals(model.Formulario.TipoMovimiento, "E", StringComparison.OrdinalIgnoreCase)
            ? operacionesEgreso
            : operacionesIngreso;
        var operacion = operaciones.FirstOrDefault(x => string.Equals(x.IdOpeBancaria, model.Formulario.IdOpeBancaria, StringComparison.OrdinalIgnoreCase));
        if (operacion is not null)
        {
            model.Formulario.TipoOperacionTexto = operacion.TipoOperacion;
        }

        if (cuentasPlan.Count == 0)
        {
            return;
        }

        var cuentasPorId = cuentasPlan.ToDictionary(x => x.IdPlanCuenta);
        foreach (var detalle in model.Formulario.Detalles)
        {
            if (detalle.IdPlanCuenta.HasValue && cuentasPorId.TryGetValue(detalle.IdPlanCuenta.Value, out var cuenta))
            {
                detalle.CuentaTexto = $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}";
                detalle.RequiereCentroCostoCuenta = cuenta.RequiereCentroCosto;
            }
        }
    }

    private static void NormalizarFormulario(CajaBancoFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue
                     || x.IdPersona.HasValue
                     || !string.IsNullOrWhiteSpace(x.ModuloOperacionComprobante)
                     || x.IdRegistroComprobante.HasValue
                     || !string.IsNullOrWhiteSpace(x.PersonaTexto)
                     || !string.IsNullOrWhiteSpace(x.GlosaDetalle)
                     || !string.IsNullOrWhiteSpace(x.CodigoCentroCosto)
                     || !string.IsNullOrWhiteSpace(x.NumeroDocumento)
                     || !string.IsNullOrWhiteSpace(x.TipoDocumento)
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

    private void ValidarFormulario(
        CajaBancoFormViewModel formulario,
        IReadOnlyDictionary<int, BancoConfiguracionEmpresaDto> cuentasCorrientesValidas,
        ISet<string> operacionesValidas,
        IReadOnlyDictionary<int, PlanCuentaDto> cuentasMovimiento,
        IReadOnlyDictionary<string, CentroCostoDto> centrosCostoActivos,
        IReadOnlyDictionary<int, PersonaDetalleDto> personasDetalle)
    {
        if (formulario.IdBancoConfiguracionEmpresa is null or <= 0)
        {
            ModelState.AddModelError("Formulario.IdBancoConfiguracionEmpresa", "Seleccione una cuenta corriente.");
        }
        else if (!cuentasCorrientesValidas.ContainsKey(formulario.IdBancoConfiguracionEmpresa.Value))
        {
            ModelState.AddModelError("Formulario.IdBancoConfiguracionEmpresa", "La cuenta corriente seleccionada no existe o no esta activa para la empresa.");
        }

        if (string.IsNullOrWhiteSpace(formulario.TipoMovimiento) || (formulario.TipoMovimiento != "I" && formulario.TipoMovimiento != "E"))
        {
            ModelState.AddModelError("Formulario.TipoMovimiento", "Seleccione si el movimiento es ingreso o egreso.");
        }

        if (string.IsNullOrWhiteSpace(formulario.IdOpeBancaria))
        {
            ModelState.AddModelError("Formulario.IdOpeBancaria", "Seleccione una operacion bancaria.");
        }
        else if (!operacionesValidas.Contains(formulario.IdOpeBancaria.Trim().ToUpperInvariant()))
        {
            ModelState.AddModelError("Formulario.IdOpeBancaria", "La operacion bancaria no corresponde al tipo de movimiento seleccionado.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Glosa))
        {
            ModelState.AddModelError("Formulario.Glosa", "Ingrese la glosa del movimiento.");
        }

        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos una linea en el detalle del movimiento.");
            return;
        }

        decimal totalDebe = 0;
        decimal totalHaber = 0;

        foreach (var i in Enumerable.Range(0, formulario.Detalles.Count))
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (!detalle.IdPlanCuenta.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione una cuenta.");
                continue;
            }

            if (!cuentasMovimiento.TryGetValue(detalle.IdPlanCuenta.Value, out var cuenta))
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "La cuenta seleccionada no esta activa o no acepta movimiento.");
                continue;
            }

            if (detalle.IdPersona.HasValue && detalle.IdPersona.Value > 0)
            {
                if (!personasDetalle.TryGetValue(detalle.IdPersona.Value, out var personaDetalle))
                {
                    ModelState.AddModelError($"{prefijo}.IdPersona", "La persona seleccionada no existe o no pertenece a la empresa activa.");
                }
                else
                {
                    detalle.PersonaTexto = personaDetalle.NombreCompleto;
                    detalle.NumeroDocumento = personaDetalle.NumeroDocumento;
                }
            }
            else
            {
                detalle.PersonaTexto = string.Empty;
                detalle.NumeroDocumento = string.IsNullOrWhiteSpace(detalle.NumeroDocumento) ? null : detalle.NumeroDocumento.Trim();
            }

            detalle.ModuloOperacionComprobante = string.IsNullOrWhiteSpace(detalle.ModuloOperacionComprobante)
                ? null
                : detalle.ModuloOperacionComprobante.Trim().ToUpperInvariant();

            detalle.RequiereCentroCostoCuenta = cuenta.RequiereCentroCosto;

            if (!string.IsNullOrWhiteSpace(detalle.CodigoCentroCosto)
                && !centrosCostoActivos.ContainsKey(detalle.CodigoCentroCosto.Trim()))
            {
                ModelState.AddModelError($"{prefijo}.CodigoCentroCosto", "El centro de costo ingresado no existe o no esta activo para la empresa.");
            }

            if (string.IsNullOrWhiteSpace(detalle.GlosaDetalle))
            {
                ModelState.AddModelError($"{prefijo}.GlosaDetalle", "Ingrese la glosa del detalle.");
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

            var comprobanteValido = detalle.ModuloOperacionComprobante is "COM" or "VEN" or "DET" or "PER" or "R4T"
                && detalle.IdRegistroComprobante.HasValue
                && detalle.IdRegistroComprobante.Value > 0;

            if (!comprobanteValido)
            {
                detalle.ModuloOperacionComprobante = null;
                detalle.IdRegistroComprobante = null;
            }

            totalDebe += detalle.Debe;
            totalHaber += detalle.Haber;
        }

        if (formulario.ImporteTotal <= 0)
        {
            ModelState.AddModelError("Formulario.ImporteTotal", "Ingrese un importe total mayor a cero.");
        }

        var totalDetalleBase = string.Equals(formulario.TipoMovimiento, "E", StringComparison.OrdinalIgnoreCase)
            ? totalDebe - totalHaber
            : totalHaber - totalDebe;
        var totalDetalle = Math.Abs(totalDetalleBase);
        var diferencia = Math.Abs(formulario.ImporteTotal - totalDetalle);

        if (diferencia >= 0.005m)
        {
            ModelState.AddModelError(string.Empty, "No puede guardar mientras exista diferencia entre Total Operacion y Total Detalle.");
        }

        if (string.Equals(formulario.TipoMovimiento, "I", StringComparison.OrdinalIgnoreCase) && totalHaber <= totalDebe)
        {
            ModelState.AddModelError(string.Empty, "En ingresos, el detalle debe tener mayor Haber que Debe para compensar la cuenta bancaria.");
        }

        if (string.Equals(formulario.TipoMovimiento, "E", StringComparison.OrdinalIgnoreCase) && totalDebe <= totalHaber)
        {
            ModelState.AddModelError(string.Empty, "En egresos, el detalle debe tener mayor Debe que Haber para compensar la cuenta bancaria.");
        }
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

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var hoy = DateTime.Today;
        return (anio ?? (short)hoy.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)hoy.Month);
    }

    private static DateOnly ParsePeriodo(string periodo)
    {
        if (periodo.Length == 6
            && int.TryParse(periodo[..4], out var anio)
            && int.TryParse(periodo[4..], out var mes)
            && mes is >= 1 and <= 12)
        {
            return new DateOnly(anio, mes, 1);
        }

        return DateOnly.FromDateTime(DateTime.Today);
    }

    private static List<int> ConstruirAnios(short anioSeleccionado)
    {
        return Enumerable.Range(anioSeleccionado - 5, 11).ToList();
    }

    private static List<MesOpcionViewModel> ConstruirMeses()
    {
        return Enumerable.Range(1, 12)
            .Select(x => new MesOpcionViewModel
            {
                Valor = (byte)x,
                Nombre = new DateTime(2000, x, 1).ToString("MMMM")
            })
            .ToList();
    }

    private static object? ResolverCuentaContableComprobante(
        ComprobanteSaldoAyudaDto comprobante,
        IReadOnlyDictionary<string, TipoComprobanteDto> tipoComprobantePorCodigo,
        IReadOnlyDictionary<int, ConfiguracionDocumentoEmpresaDto> configuracionDocumentoPorId,
        (int idPlanCuenta, string cuentaTexto)? cuentaDetraccion,
        (int idPlanCuenta, string cuentaTexto)? cuentaPercepcion,
        (int idPlanCuenta, string cuentaTexto)? cuentaRetencion)
    {
        if (string.Equals(comprobante.ModuloOperacion, "DET", StringComparison.OrdinalIgnoreCase))
        {
            if (!cuentaDetraccion.HasValue)
            {
                return null;
            }

            return new
            {
                idPlanCuenta = cuentaDetraccion.Value.idPlanCuenta,
                cuentaTexto = cuentaDetraccion.Value.cuentaTexto
            };
        }

        if (string.Equals(comprobante.ModuloOperacion, "PER", StringComparison.OrdinalIgnoreCase))
        {
            if (!cuentaPercepcion.HasValue)
            {
                return null;
            }

            return new
            {
                idPlanCuenta = cuentaPercepcion.Value.idPlanCuenta,
                cuentaTexto = cuentaPercepcion.Value.cuentaTexto
            };
        }

        if (string.Equals(comprobante.ModuloOperacion, "R4T", StringComparison.OrdinalIgnoreCase))
        {
            if (!cuentaRetencion.HasValue)
            {
                return null;
            }

            return new
            {
                idPlanCuenta = cuentaRetencion.Value.idPlanCuenta,
                cuentaTexto = cuentaRetencion.Value.cuentaTexto
            };
        }

        if (!tipoComprobantePorCodigo.TryGetValue(comprobante.TipoComprobante, out var tipoComprobante))
        {
            return null;
        }

        if (!configuracionDocumentoPorId.TryGetValue(tipoComprobante.IdTipoComprobante, out var configuracionDocumento))
        {
            return null;
        }

        var esDolares = string.Equals(comprobante.CodigoMoneda, "USD", StringComparison.OrdinalIgnoreCase);
        var esCompra = string.Equals(comprobante.ModuloOperacion, "COM", StringComparison.OrdinalIgnoreCase);

        var idPlanCuenta = esCompra
            ? (esDolares ? configuracionDocumento.IdCuentaCompraDolares : configuracionDocumento.IdCuentaCompraSoles)
            : (esDolares ? configuracionDocumento.IdCuentaVentaDolares : configuracionDocumento.IdCuentaVentaSoles);

        var cuentaTexto = esCompra
            ? (esDolares ? configuracionDocumento.CuentaCompraDolaresTexto : configuracionDocumento.CuentaCompraSolesTexto)
            : (esDolares ? configuracionDocumento.CuentaVentaDolaresTexto : configuracionDocumento.CuentaVentaSolesTexto);

        if (!idPlanCuenta.HasValue || idPlanCuenta.Value <= 0 || string.IsNullOrWhiteSpace(cuentaTexto))
        {
            return null;
        }

        return new
        {
            idPlanCuenta = idPlanCuenta.Value,
            cuentaTexto
        };
    }

    private async Task<(int idPlanCuenta, string cuentaTexto)?> ObtenerCuentaDetraccionAsync(int idEmpresa, CancellationToken cancellationToken)
    {
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, null, "CTADETRACCION", 1, 20, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x =>
            x.Activo &&
            string.Equals(x.CodigoParametro, "CTADETRACCION", StringComparison.OrdinalIgnoreCase));

        if (parametro is null || string.IsNullOrWhiteSpace(parametro.ValorParametro))
        {
            return null;
        }

        var valorParametro = parametro.ValorParametro.Trim();
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(idEmpresa, true, cancellationToken);

        var cuenta = cuentas.FirstOrDefault(x =>
            string.Equals(x.CodigoCuenta, valorParametro, StringComparison.OrdinalIgnoreCase)
            && x.Estado
            && x.AceptaMovimiento);

        if (cuenta is null && int.TryParse(valorParametro, out var idPlanCuenta) && idPlanCuenta > 0)
        {
            cuenta = cuentas.FirstOrDefault(x => x.IdPlanCuenta == idPlanCuenta && x.Estado && x.AceptaMovimiento);
        }

        if (cuenta is null)
        {
            return null;
        }

        return (cuenta.IdPlanCuenta, $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}");
    }

    private static (int idPlanCuenta, string cuentaTexto)? ObtenerCuentaRetencion(ConfiguracionContableEmpresaDto configuracionContable)
    {
        var impuesto = configuracionContable.Impuestos.FirstOrDefault(x =>
            x.Activo
            && string.Equals(x.CodigoSunat, "R4TA", StringComparison.OrdinalIgnoreCase)
            && x.IdPlanCuenta.HasValue
            && x.IdPlanCuenta.Value > 0
            && !string.IsNullOrWhiteSpace(x.CuentaTexto));

        if (impuesto is null)
        {
            return null;
        }

        return (impuesto.IdPlanCuenta!.Value, impuesto.CuentaTexto);
    }

    private async Task<(int idPlanCuenta, string cuentaTexto)?> ObtenerCuentaPercepcionAsync(int idEmpresa, CancellationToken cancellationToken)
    {
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, null, "CTADEPERCEPCION", 1, 20, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x =>
            x.Activo &&
            string.Equals(x.CodigoParametro, "CTADEPERCEPCION", StringComparison.OrdinalIgnoreCase));

        if (parametro is null || string.IsNullOrWhiteSpace(parametro.ValorParametro))
        {
            return null;
        }

        var valorParametro = parametro.ValorParametro.Trim();
        var cuentas = await planCuentaRepository.ListarPorEmpresaAsync(idEmpresa, true, cancellationToken);

        var cuenta = cuentas.FirstOrDefault(x =>
            string.Equals(x.CodigoCuenta, valorParametro, StringComparison.OrdinalIgnoreCase)
            && x.Estado
            && x.AceptaMovimiento);

        if (cuenta is null && int.TryParse(valorParametro, out var idPlanCuenta) && idPlanCuenta > 0)
        {
            cuenta = cuentas.FirstOrDefault(x => x.IdPlanCuenta == idPlanCuenta && x.Estado && x.AceptaMovimiento);
        }

        if (cuenta is null)
        {
            return null;
        }

        return (cuenta.IdPlanCuenta, $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}");
    }
}
