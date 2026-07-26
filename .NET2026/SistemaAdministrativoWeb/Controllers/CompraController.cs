using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("COMPRAS")]
public class CompraController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPeriodoContableService periodoContableService,
    ICompraRepository compraRepository,
    IXmlProvisionImportService xmlProvisionImportService,
    IMigoPadronApiClient migoPadronApiClient,
    IProveedorRepository proveedorRepository,
    IPersonaRepository personaRepository,
    IConfiguracionContabilizacionRepository configuracionRepository,
    IAsientoPreviewService asientoPreviewService,
    IPlanCuentaRepository planCuentaRepository,
    ITipoAfectacionIgvRepository tipoAfectacionIgvRepository,
    IDetraccionSunatRepository detraccionSunatRepository,
    ITipoPercepcionRepository tipoPercepcionRepository,
    IParametroEmpresaRepository parametroEmpresaRepository,
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository,
    IPlanFeatureService planFeatureService) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;
    private const string CodigoDocumentoRucSunat = "6";
    private const string CodigoAfectacionGravadoOnerosa = "10";
    private const string CodigoReciboHonorarios = "02";
    private const string CodigoParametroPorcentajeRetencion4ta = "PORCRETEN4TA";
    private const decimal TasaIgv = 0.18m;
    private static readonly HashSet<string> TiposComprobanteValidacionCpe = new(StringComparer.OrdinalIgnoreCase)
    {
        "01",
        "03",
        "07",
        "08",
        "02"
    };

    [HttpGet]
    public async Task<IActionResult> Index(short? anio = null, byte? mes = null, string? textoBusqueda = null, string? tipoComprobante = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var proveedores = (await proveedorRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "COM")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(true, false, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var detraccionesSunat = (await detraccionSunatRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var tiposPercepcion = (await tipoPercepcionRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.Codigo)
            .ToList();
        var porcentajeRetencionRenta4ta = await ObtenerPorcentajeRetencionRenta4taAsync(empresaId, cancellationToken);
        var tipoComprobanteFiltro = string.IsNullOrWhiteSpace(tipoComprobante) ? null : tipoComprobante.Trim().ToUpperInvariant();
        var compras = await compraRepository.ListarPaginadoPorEmpresaAsync(empresaId, anioTrabajo, mesTrabajo, textoBusqueda, tipoComprobanteFiltro, pagina, TamanoPagina, cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            tipoComprobanteFiltro,
            proveedores,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            detraccionesSunat,
            tiposPercepcion,
            porcentajeRetencionRenta4ta,
            compras.Items,
            null);
        model.TotalCompras = compras.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = compras.TotalRecords
        };
        model.PuedeValidarCpe = await planFeatureService.IsEnabledAsync(
            User,
            PlanFeature.CpeValidation,
            cancellationToken);

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(string? periodo = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = DescomponerPeriodo(periodoTrabajo);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        return await CargarFormularioAsync(periodo, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> CargaMasiva(string? periodo = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = (short.Parse(periodoTrabajo[..4]), byte.Parse(periodoTrabajo[4..]));
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        ViewData["AdminShell"] = true;

        return View(new CargaMasivaXmlViewModel
        {
            Titulo = "Carga masiva de compras",
            Subtitulo = "Importa XML SUNAT para registrar provisiones en revision y completar luego la cuenta contable.",
            Modulo = "COM",
            PeriodoConsulta = periodoTrabajo,
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> CargaMasiva(string? periodo, List<IFormFile>? archivos, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = (short.Parse(periodoTrabajo[..4]), byte.Parse(periodoTrabajo[4..]));
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        ViewData["AdminShell"] = true;
        var model = new CargaMasivaXmlViewModel
        {
            Titulo = "Carga masiva de compras",
            Subtitulo = "Importa XML SUNAT para registrar provisiones en revision y completar luego la cuenta contable.",
            Modulo = "COM",
            PeriodoConsulta = periodoTrabajo,
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo
        };

        if (archivos is null || archivos.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Seleccione al menos un archivo XML.");
            return View(model);
        }

        var archivosImportacion = new List<ImportacionXmlArchivoRequest>();
        foreach (var archivo in archivos.Where(x => x is not null && x.Length > 0))
        {
            if (!string.Equals(Path.GetExtension(archivo.FileName), ".xml", StringComparison.OrdinalIgnoreCase))
            {
                ModelState.AddModelError(string.Empty, $"El archivo {archivo.FileName} no es XML.");
                continue;
            }

            await using var stream = new MemoryStream();
            await archivo.CopyToAsync(stream, cancellationToken);
            archivosImportacion.Add(new ImportacionXmlArchivoRequest
            {
                NombreArchivo = archivo.FileName,
                Contenido = stream.ToArray()
            });
        }

        if (!ModelState.IsValid)
        {
            return View(model);
        }

        try
        {
            var resultado = await xmlProvisionImportService.ImportarComprasAsync(
                currentCompanyAccessor.EmpresaId.Value,
                archivosImportacion,
                User.Identity?.Name,
                cancellationToken);

            model.Resultados = resultado.Items;
            if (resultado.Items.Any(x => x.Importado))
            {
                TempData["CompraOk"] = $"Se importaron {resultado.Items.Count(x => x.Importado)} compra(s) desde XML.";
            }
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
        }

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idCompra, string? periodo = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = DescomponerPeriodo(periodoTrabajo);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        return await CargarFormularioAsync(periodo, idCompra, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idCompra, short? anio = null, byte? mes = null, string? textoBusqueda = null, string? tipoComprobante = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
        }

        try
        {
            await compraRepository.EliminarAsync(idCompra, currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            TempData["CompraOk"] = "Compra eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["CompraError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [Authorize(Policy = PlanFeaturePolicies.CpeValidation)]
    public async Task<IActionResult> ValidarCpe(int idCompra, short? anio = null, byte? mes = null, string? textoBusqueda = null, string? tipoComprobante = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["CompraError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
        }

        try
        {
            var compra = await compraRepository.ObtenerAsync(idCompra, cancellationToken);
            if (compra is null || compra.IdEmpresa != currentCompanyAccessor.EmpresaId.Value)
            {
                TempData["CompraError"] = "No se encontro la compra seleccionada.";
                return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
            }

            var codigoTipoComprobante = (compra.TipoComprobante ?? string.Empty).Trim().ToUpperInvariant();
            if (!TiposComprobanteValidacionCpe.Contains(codigoTipoComprobante))
            {
                TempData["CompraError"] = "El comprobante seleccionado no admite validacion CPE.";
                return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
            }

            if (string.IsNullOrWhiteSpace(compra.NumeroDocumentoPersona) || compra.NumeroDocumentoPersona.Length != 11)
            {
                TempData["CompraError"] = "La compra no tiene un RUC emisor valido para consultar el CPE.";
                return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
            }

            var resultado = await migoPadronApiClient.ValidarCpeAsync(new MigoCpeRequestDto
            {
                RucEmisor = compra.NumeroDocumentoPersona,
                TipoComprobante = codigoTipoComprobante,
                Serie = compra.Serie,
                Numero = compra.Numero,
                FechaEmision = compra.FechaEmision,
                Monto = compra.ImporteTotal
            }, cancellationToken);

            if (resultado is null || !resultado.Success)
            {
                TempData["CompraError"] = string.IsNullOrWhiteSpace(resultado?.Observaciones)
                    ? "La API no devolvio una respuesta valida al consultar el CPE."
                    : resultado!.Observaciones;
                return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
            }

            var estadoValidacion = DescribirEstadoComprobante(resultado.EstadoComprobante);
            var mensajeValidacion = ConstruirMensajeValidacionCpe(resultado);

            await compraRepository.GuardarValidacionCpeAsync(new GuardarCompraValidacionCpeRequest
            {
                IdCompra = compra.IdCompra,
                IdEmpresa = compra.IdEmpresa,
                FechaValidacionCpe = DateTime.Now,
                EstadoValidacionCpe = estadoValidacion,
                MensajeValidacionCpe = mensajeValidacion,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["CompraOk"] = $"Validacion CPE registrada: {estadoValidacion}.";
        }
        catch (Exception ex)
        {
            TempData["CompraError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, tipoComprobante, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(CompraFormViewModel formulario, string? periodo = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .ToList();

        var porcentajeRetencionRenta4ta = await ObtenerPorcentajeRetencionRenta4taAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
        NormalizarFormulario(formulario, tiposAfectacionIgv, porcentajeRetencionRenta4ta);
        ValidarFormulario(formulario);

        var periodoTrabajo = string.IsNullOrWhiteSpace(periodo)
            ? $"{formulario.FechaContabilizacion.Year:0000}{formulario.FechaContabilizacion.Month:00}"
            : NormalizarPeriodo(periodo);
        var (anioTrabajo, mesTrabajo) = DescomponerPeriodo(periodoTrabajo);
        if (await periodoContableService.EstaCerradoAsync(
                currentCompanyAccessor.EmpresaId.Value,
                (short)formulario.FechaContabilizacion.Year,
                (byte)formulario.FechaContabilizacion.Month,
                cancellationToken))
        {
            ModelState.AddModelError(
                string.Empty,
                periodoContableService.ConstruirMensajeBloqueo(
                    (short)formulario.FechaContabilizacion.Year,
                    (byte)formulario.FechaContabilizacion.Month));
        }

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var result = await compraRepository.GuardarAsync(new GuardarCompraRequest
            {
                IdCompra = formulario.IdCompra,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdProveedor = formulario.IdProveedor!.Value,
                IdConfiguracionContabilizacion = formulario.IdConfiguracionContabilizacion!.Value,
                FechaEmision = formulario.FechaEmision,
                FechaContabilizacion = formulario.FechaContabilizacion,
                TipoComprobante = formulario.TipoComprobante.Trim().ToUpperInvariant(),
                Serie = formulario.Serie.Trim().ToUpperInvariant(),
                Numero = formulario.Numero.Trim().ToUpperInvariant(),
                IdMoneda = formulario.IdMoneda!.Value,
                TipoCambio = formulario.TipoCambio,
                BaseImponible = formulario.BaseImponible,
                TotalExonerado = formulario.TotalExonerado,
                TotalInafecto = formulario.TotalInafecto,
                Icbper = formulario.Icbper,
                Igv = formulario.Igv,
                Isc = formulario.Isc,
                OtrosTributos = formulario.OtrosTributos,
                Redondeo = formulario.Redondeo,
                ImporteTotal = formulario.ImporteTotal,
                ExoneracionRenta4ta = formulario.ExoneracionRenta4ta,
                Retencion = formulario.Retencion,
                TieneDetraccion = formulario.TieneDetraccion,
                IdDetraccionSunat = formulario.TieneDetraccion ? formulario.IdDetraccionSunat : null,
                ImporteDetraccion = formulario.TieneDetraccion ? formulario.ImporteDetraccion : 0m,
                TienePercepcion = formulario.TienePercepcion,
                IdTipoPercepcion = formulario.TienePercepcion ? formulario.IdTipoPercepcion : null,
                BasePercepcion = formulario.TienePercepcion ? formulario.BasePercepcion : 0m,
                ImportePercepcion = formulario.TienePercepcion ? formulario.ImportePercepcion : 0m,
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarCompraDetalleRequest
                    {
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        IdTipoAfectacionIGV = x.IdTipoAfectacionIGV!.Value,
                        Descripcion = x.Descripcion.Trim(),
                        Cantidad = x.Cantidad,
                        ValorUnitario = x.ValorUnitario,
                        ImporteBruto = x.ImporteBruto
                    })
                    .ToList()
            }, cancellationToken);

            TempData["CompraOk"] = formulario.TieneDetraccion && result.IdAsientoDetraccion.HasValue
                ? $"Compra provisionada correctamente. Asiento compra: {(result.IdAsiento.HasValue ? result.IdAsiento.Value.ToString() : "sin asiento")} | Asiento detracción: {result.IdAsientoDetraccion.Value}."
                : $"Compra provisionada correctamente. Asiento vinculado: {(result.IdAsiento.HasValue ? result.IdAsiento.Value.ToString() : "sin asiento")}.";
            var mensajesAsiento = new List<string>
            {
                $"Asiento compra: {(result.IdAsiento.HasValue ? result.IdAsiento.Value.ToString() : "sin asiento")}"
            };

            if (formulario.TieneDetraccion && result.IdAsientoDetraccion.HasValue)
            {
                mensajesAsiento.Add($"Asiento detraccion: {result.IdAsientoDetraccion.Value}");
            }

            if (formulario.TienePercepcion && result.IdAsientoPercepcion.HasValue)
            {
                mensajesAsiento.Add($"Asiento percepcion: {result.IdAsientoPercepcion.Value}");
            }

            TempData["CompraOk"] = $"Compra provisionada correctamente. {string.Join(" | ", mensajesAsiento)}.";
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
    public async Task<IActionResult> BuscarProveedores(string? buscar = null, int? proveedorId = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var criterio = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var proveedores = await proveedorRepository.ListarActivosPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);

        var items = proveedores
            .Where(x =>
                proveedorId.HasValue && x.IdProveedor == proveedorId.Value
                || criterio is null
                || x.NombreCompleto.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || x.NumeroDocumento.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || x.CodigoProveedor.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || (!string.IsNullOrWhiteSpace(x.Contacto) && x.Contacto.Contains(criterio, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(x => x.NombreCompleto)
            .Take(30)
            .Select(x => new
            {
                value = x.IdProveedor,
                text = $"{x.NombreCompleto} ({x.NumeroDocumento})",
                tipoDocumento = x.TipoDocumento,
                numeroDocumento = x.NumeroDocumento,
                nombre = x.NombreCompleto,
                numero = x.Contacto ?? x.Telefono ?? string.Empty,
                correo = x.CorreoElectronico ?? string.Empty
            })
            .ToList();

        return Json(new { ok = true, items });
    }

    [HttpPost]
    public async Task<IActionResult> CrearProveedorRapido([FromBody] RegistroRapidoPersonaRequestViewModel request, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
            {
                return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
            }

            var tipoPersona = (request.TipoPersona ?? string.Empty).Trim().ToUpperInvariant();
            var tipoDocumento = (request.TipoDocumento ?? string.Empty).Trim();
            var numeroDocumento = (request.NumeroDocumento ?? string.Empty).Trim();
            var razonSocial = (request.RazonSocial ?? string.Empty).Trim();
            var nombres = (request.Nombres ?? string.Empty).Trim();
            var apellidos = (request.Apellidos ?? string.Empty).Trim();
            var telefono = string.IsNullOrWhiteSpace(request.Telefono) ? null : request.Telefono.Trim();
            var correo = string.IsNullOrWhiteSpace(request.Correo) ? null : request.Correo.Trim();
            var esJuridica = string.Equals(tipoPersona, "J", StringComparison.OrdinalIgnoreCase);

            if (!string.Equals(tipoPersona, "N", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(tipoPersona, "J", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { ok = false, mensaje = "Seleccione el tipo de persona." });
            }

            if (string.IsNullOrWhiteSpace(tipoDocumento))
            {
                return BadRequest(new { ok = false, mensaje = "Seleccione el tipo de documento." });
            }

            if (string.IsNullOrWhiteSpace(numeroDocumento))
            {
                return BadRequest(new { ok = false, mensaje = "Ingrese el numero de documento." });
            }

            if (esJuridica)
            {
                if (string.IsNullOrWhiteSpace(razonSocial))
                {
                    return BadRequest(new { ok = false, mensaje = "Ingrese la razon social del proveedor." });
                }
            }
            else if (string.IsNullOrWhiteSpace(nombres) || string.IsNullOrWhiteSpace(apellidos))
            {
                return BadRequest(new { ok = false, mensaje = "Ingrese nombres y apellidos del proveedor." });
            }

            await personaRepository.GuardarAsync(new GuardarPersonaRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                TipoPersona = tipoPersona,
                TipoDocumento = tipoDocumento,
                NumeroDocumento = numeroDocumento,
                ApellidoPaterno = esJuridica ? null : apellidos,
                ApellidoMaterno = null,
                Nombres = esJuridica ? null : nombres,
                RazonSocial = esJuridica ? razonSocial : null,
                CorreoElectronico = correo,
                Telefono = telefono,
                Direccion = null,
                CodigoUbigeo = null,
                EsCliente = false,
                EsProveedor = true,
                Estado = true,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            var proveedores = await proveedorRepository.ListarActivosPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var proveedor = proveedores.FirstOrDefault(x =>
                string.Equals(x.TipoDocumento, tipoDocumento, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.NumeroDocumento, numeroDocumento, StringComparison.OrdinalIgnoreCase));

            if (proveedor is null)
            {
                return BadRequest(new { ok = false, mensaje = "El proveedor fue registrado, pero no pudo recuperarse para la seleccion." });
            }

            return Json(new
            {
                ok = true,
                proveedorId = proveedor.IdProveedor,
                proveedorTexto = $"{proveedor.NombreCompleto} ({proveedor.NumeroDocumento})",
                tipoDocumento = proveedor.TipoDocumento,
                numeroDocumento = proveedor.NumeroDocumento,
                nombre = proveedor.NombreCompleto,
                numero = proveedor.Contacto ?? proveedor.Telefono ?? string.Empty,
                correo = proveedor.CorreoElectronico ?? string.Empty
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> PrevisualizarAsiento([FromBody] AsientoPreviewRequestViewModel request, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        if (!request.IdConfiguracionContabilizacion.HasValue || request.IdConfiguracionContabilizacion.Value <= 0)
        {
            return BadRequest(new { ok = false, mensaje = "Seleccione una configuracion contable." });
        }

        try
        {
            var preview = await asientoPreviewService.PrevisualizarAsync(currentCompanyAccessor.EmpresaId.Value, new AsientoPreviewRequest
            {
                ModuloOperacion = "COM",
                IdConfiguracionContabilizacion = request.IdConfiguracionContabilizacion.Value,
                FechaContabilizacion = request.FechaContabilizacion,
                TipoComprobante = request.TipoComprobante ?? string.Empty,
                Serie = request.Serie ?? string.Empty,
                Numero = request.Numero ?? string.Empty,
                BaseImponible = request.BaseImponible,
                Igv = request.Igv,
                Isc = request.Isc,
                OtrosTributos = request.OtrosTributos,
                Redondeo = request.Redondeo,
                ImporteTotal = request.ImporteTotal,
                Detalles = request.Detalles
                    .Select(x => new AsientoPreviewDetalleRequest
                    {
                        Item = x.Item,
                        Descripcion = x.Descripcion,
                        Cantidad = x.Cantidad,
                        ValorUnitario = x.ValorUnitario,
                        ImporteBruto = x.ImporteBruto
                    })
                    .ToList()
            }, cancellationToken);

            return Json(new { ok = true, preview });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(string? periodo, int? idCompra, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var periodoTrabajo = NormalizarPeriodo(periodo);
        var proveedores = (await proveedorRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "COM")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(true, false, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var detraccionesSunat = (await detraccionSunatRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var tiposPercepcion = (await tipoPercepcionRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.Codigo)
            .ToList();
        var porcentajeRetencionRenta4ta = await ObtenerPorcentajeRetencionRenta4taAsync(empresaId, cancellationToken);
        var compras = await compraRepository.ListarPorEmpresaAsync(empresaId, periodoTrabajo, cancellationToken);
        var compraEditar = idCompra.HasValue
            ? await compraRepository.ObtenerAsync(idCompra.Value, cancellationToken)
            : null;

        if (compraEditar is not null && compraEditar.IdEmpresa != empresaId)
        {
            compraEditar = null;
        }

        return View("Formulario", ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            short.Parse(periodoTrabajo[..4]),
            byte.Parse(periodoTrabajo[4..]),
            null,
            null,
            proveedores,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            detraccionesSunat,
            tiposPercepcion,
            porcentajeRetencionRenta4ta,
            compras,
            compraEditar));
    }

    private async Task<CompraIndexViewModel> ConstruirViewModelErrorAsync(int empresaId, string periodo, CompraFormViewModel formulario, CancellationToken cancellationToken)
    {
        var proveedores = (await proveedorRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "COM")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(true, false, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var detraccionesSunat = (await detraccionSunatRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var tiposPercepcion = (await tipoPercepcionRepository.ListarActivasAsync(cancellationToken))
            .OrderBy(x => x.Codigo)
            .ToList();
        var porcentajeRetencionRenta4ta = await ObtenerPorcentajeRetencionRenta4taAsync(empresaId, cancellationToken);
        var compras = await compraRepository.ListarPorEmpresaAsync(empresaId, periodo, cancellationToken);
        var cuentasMovimientoLookup = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .ToDictionary(x => x.IdPlanCuenta);

        foreach (var detalle in formulario.Detalles.Where(x => x.IdPlanCuenta.HasValue && string.IsNullOrWhiteSpace(x.CuentaTexto)))
        {
            if (detalle.IdPlanCuenta.HasValue
                && cuentasMovimientoLookup.TryGetValue(detalle.IdPlanCuenta.Value, out var cuenta))
            {
                detalle.CuentaTexto = $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}";
            }
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodo,
            short.Parse(periodo[..4]),
            byte.Parse(periodo[4..]),
            null,
            null,
            proveedores,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            detraccionesSunat,
            tiposPercepcion,
            porcentajeRetencionRenta4ta,
            compras,
            null);

        model.Formulario = formulario;
        return model;
    }

    private static void NormalizarFormulario(
        CompraFormViewModel formulario,
        IReadOnlyCollection<TipoAfectacionIgvDto> tiposAfectacionIgv,
        decimal porcentajeRetencionRenta4ta)
    {
        formulario.Serie = NormalizarSerieDocumento(formulario.Serie);
        formulario.Numero = NormalizarNumeroDocumento(formulario.Numero);

        var idAfectacionGravada = tiposAfectacionIgv
            .FirstOrDefault(x => x.CodigoSunat == CodigoAfectacionGravadoOnerosa)
            ?.IdTipoAfectacionIGV ?? 1;

        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue || !string.IsNullOrWhiteSpace(x.Descripcion) || x.ImporteBruto > 0 || x.ValorUnitario > 0 || x.Cantidad > 0)
            .Select((x, index) =>
            {
                x.Item = (short)(index + 1);
                x.IdTipoAfectacionIGV ??= idAfectacionGravada;
                x.ImporteBruto = decimal.Round(x.Cantidad * x.ValorUnitario, 2);
                return x;
            })
            .ToList();

        var codigosAfectacion = tiposAfectacionIgv.ToDictionary(x => x.IdTipoAfectacionIGV, x => x.CodigoSunat);
        formulario.BaseImponible = decimal.Round(formulario.Detalles.Sum(x => x.ImporteBruto), 2);
        formulario.TotalExonerado = decimal.Round(
            formulario.Detalles
                .Where(x => x.IdTipoAfectacionIGV.HasValue
                    && codigosAfectacion.TryGetValue(x.IdTipoAfectacionIGV.Value, out var codigo)
                    && EsAfectacionExonerada(codigo))
                .Sum(x => x.ImporteBruto),
            2);
        formulario.TotalInafecto = decimal.Round(
            formulario.Detalles
                .Where(x => x.IdTipoAfectacionIGV.HasValue
                    && codigosAfectacion.TryGetValue(x.IdTipoAfectacionIGV.Value, out var codigo)
                    && EsAfectacionInafecta(codigo))
                .Sum(x => x.ImporteBruto),
            2);
        formulario.Igv = decimal.Round(
            formulario.Detalles
                .Where(x => x.IdTipoAfectacionIGV.HasValue
                    && codigosAfectacion.TryGetValue(x.IdTipoAfectacionIGV.Value, out var codigo)
                    && EsAfectacionGravada(codigo))
                .Sum(x => x.ImporteBruto) * TasaIgv,
            2);
        var esReciboHonorarios = string.Equals(formulario.TipoComprobante, CodigoReciboHonorarios, StringComparison.OrdinalIgnoreCase);
        formulario.PorcentajeRetencion = esReciboHonorarios ? porcentajeRetencionRenta4ta : 0m;
        formulario.ExoneracionRenta4ta = esReciboHonorarios && formulario.ExoneracionRenta4ta;
        formulario.Retencion = esReciboHonorarios && !formulario.ExoneracionRenta4ta
            ? decimal.Round(formulario.BaseImponible * (formulario.PorcentajeRetencion / 100m), 2)
            : 0m;
        if (esReciboHonorarios)
        {
            formulario.Igv = 0m;
        }
        formulario.Icbper = 0m;
        formulario.Isc = 0m;
        formulario.OtrosTributos = 0m;
        formulario.Redondeo = 0m;
        formulario.ImporteTotal = esReciboHonorarios
            ? decimal.Round(formulario.BaseImponible - formulario.Retencion, 2)
            : formulario.BaseImponible + formulario.Igv;
        formulario.ImporteDetraccion = formulario.TieneDetraccion
            ? decimal.Round(formulario.ImporteTotal * (formulario.PorcentajeDetraccion / 100m), 2)
            : 0m;
        formulario.BasePercepcion = formulario.TienePercepcion
            ? formulario.ImporteTotal
            : 0m;
        formulario.ImportePercepcion = formulario.TienePercepcion
            ? decimal.Round(formulario.BasePercepcion * (formulario.PorcentajePercepcion / 100m), 2)
            : 0m;
        formulario.SaldoPago = decimal.Round(formulario.ImporteTotal - formulario.ImporteDetraccion, 2);
    }

    private void ValidarFormulario(CompraFormViewModel formulario)
    {
        var esReciboHonorarios = string.Equals(formulario.TipoComprobante, CodigoReciboHonorarios, StringComparison.OrdinalIgnoreCase);

        if (!formulario.TieneDetraccion)
        {
            ModelState.Remove(nameof(formulario.IdDetraccionSunat));
            ModelState.Remove(nameof(formulario.DetraccionTexto));
            ModelState.Remove(nameof(formulario.PorcentajeDetraccion));
            ModelState.Remove(nameof(formulario.ImporteDetraccion));
        }

        if (!formulario.TienePercepcion)
        {
            ModelState.Remove(nameof(formulario.IdTipoPercepcion));
            ModelState.Remove(nameof(formulario.TipoPercepcionTexto));
            ModelState.Remove(nameof(formulario.PorcentajePercepcion));
            ModelState.Remove(nameof(formulario.BasePercepcion));
            ModelState.Remove(nameof(formulario.ImportePercepcion));
        }

        if (!esReciboHonorarios)
        {
            ModelState.Remove(nameof(formulario.ExoneracionRenta4ta));
            ModelState.Remove(nameof(formulario.PorcentajeRetencion));
            ModelState.Remove(nameof(formulario.Retencion));
        }

        if (formulario.IdProveedor.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdProveedor), "Seleccione un proveedor.");
        }

        if (formulario.IdConfiguracionContabilizacion.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdConfiguracionContabilizacion), "Seleccione una configuracion contable.");
        }

        if (formulario.IdMoneda.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdMoneda), "Seleccione la moneda.");
        }

        if (string.IsNullOrWhiteSpace(formulario.TipoComprobante))
        {
            ModelState.AddModelError(nameof(formulario.TipoComprobante), "Seleccione el tipo de comprobante.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Serie))
        {
            ModelState.AddModelError(nameof(formulario.Serie), "Ingrese la serie del documento.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Numero))
        {
            ModelState.AddModelError(nameof(formulario.Numero), "Ingrese el numero del documento.");
        }

        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos un concepto en la compra.");
        }

        decimal totalDetalle = 0;

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (string.IsNullOrWhiteSpace(detalle.Descripcion))
            {
                ModelState.AddModelError($"{prefijo}.Descripcion", "Ingrese la descripcion del concepto.");
            }

            if (!detalle.IdPlanCuenta.HasValue || detalle.IdPlanCuenta.Value <= 0)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione la cuenta contable.");
            }

            if (!detalle.IdTipoAfectacionIGV.HasValue || detalle.IdTipoAfectacionIGV.Value <= 0)
            {
                ModelState.AddModelError($"{prefijo}.IdTipoAfectacionIGV", "Seleccione el tipo de afectacion IGV.");
            }

            totalDetalle += detalle.ImporteBruto;
        }

        if (esReciboHonorarios)
        {
            if (formulario.Igv != 0)
            {
                ModelState.AddModelError(nameof(formulario.Igv), "Los recibos por honorarios no deben calcular IGV.");
            }

            if (!formulario.ExoneracionRenta4ta && formulario.PorcentajeRetencion <= 0)
            {
                ModelState.AddModelError(nameof(formulario.PorcentajeRetencion), "Configure un porcentaje valido en el parametro PORCRETEN4TA.");
            }

            var importeEsperado = decimal.Round(formulario.BaseImponible - formulario.Retencion, 2);
            if (formulario.ImporteTotal != importeEsperado)
            {
                ModelState.AddModelError(string.Empty, "El importe total del recibo por honorarios debe ser igual al subtotal menos la retencion.");
            }
        }
        else if (formulario.ImporteTotal != formulario.BaseImponible + formulario.Igv)
        {
            ModelState.AddModelError(string.Empty, "El importe total debe ser igual a la suma del subtotal e IGV.");
        }

        if (formulario.TieneDetraccion)
        {
            if (!formulario.IdDetraccionSunat.HasValue || formulario.IdDetraccionSunat.Value <= 0)
            {
                ModelState.AddModelError(nameof(formulario.IdDetraccionSunat), "Seleccione el codigo de detraccion.");
            }

            if (formulario.PorcentajeDetraccion <= 0)
            {
                ModelState.AddModelError(nameof(formulario.PorcentajeDetraccion), "La detraccion seleccionada no tiene un porcentaje valido.");
            }

            if (formulario.ImporteDetraccion <= 0)
            {
                ModelState.AddModelError(nameof(formulario.ImporteDetraccion), "El importe de detraccion debe ser mayor a cero.");
            }

            if (formulario.ImporteDetraccion >= formulario.ImporteTotal && formulario.ImporteTotal > 0)
            {
                ModelState.AddModelError(nameof(formulario.ImporteDetraccion), "La detraccion debe ser menor al importe total de la compra.");
            }
        }

        if (formulario.TienePercepcion)
        {
            if (!formulario.IdTipoPercepcion.HasValue || formulario.IdTipoPercepcion.Value <= 0)
            {
                ModelState.AddModelError(nameof(formulario.IdTipoPercepcion), "Seleccione el tipo de percepcion.");
            }

            if (formulario.PorcentajePercepcion <= 0)
            {
                ModelState.AddModelError(nameof(formulario.PorcentajePercepcion), "El tipo de percepcion seleccionado no tiene un porcentaje valido.");
            }

            if (decimal.Round(formulario.BasePercepcion, 2) != decimal.Round(formulario.ImporteTotal, 2))
            {
                ModelState.AddModelError(nameof(formulario.BasePercepcion), "La base de percepcion debe ser igual al total del comprobante incluido IGV.");
            }

            if (formulario.ImportePercepcion <= 0)
            {
                ModelState.AddModelError(nameof(formulario.ImportePercepcion), "El importe de percepcion debe ser mayor a cero.");
            }
        }

        if (formulario.BaseImponible > 0 && totalDetalle > 0 && decimal.Round(totalDetalle, 2) != decimal.Round(formulario.BaseImponible, 2))
        {
            ModelState.AddModelError(string.Empty, "La suma del detalle debe coincidir con la base imponible.");
        }
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var today = DateTime.Today;
        return (anio ?? (short)today.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)today.Month);
    }

    private static string NormalizarPeriodo(string? periodo)
    {
        if (!string.IsNullOrWhiteSpace(periodo)
            && periodo.Length == 6
            && short.TryParse(periodo[..4], out var anio)
            && byte.TryParse(periodo[4..], out var mes)
            && mes is >= 1 and <= 12)
        {
            return $"{anio:0000}{mes:00}";
        }

        var (anioActual, mesActual) = NormalizarPeriodo(null, null);
        return $"{anioActual:0000}{mesActual:00}";
    }

    private static CompraIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        string periodo,
        short anioSeleccionado,
        byte mesSeleccionado,
        string? textoBusqueda,
        string? tipoComprobanteFiltro,
        IReadOnlyCollection<ProveedorDto> proveedores,
        IReadOnlyCollection<ConfiguracionContabilizacionResumenDto> configuraciones,
        IReadOnlyCollection<OpcionCatalogoViewModel> tiposDocumentoIdentidad,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<TipoComprobanteDto> tiposComprobante,
        IReadOnlyCollection<PlanCuentaDto> cuentasMovimiento,
        IReadOnlyCollection<TipoAfectacionIgvDto> tiposAfectacionIgv,
        IReadOnlyCollection<DetraccionSunatDto> detraccionesSunat,
        IReadOnlyCollection<TipoPercepcionDto> tiposPercepcion,
        decimal porcentajeRetencionRenta4ta,
        IReadOnlyCollection<CompraResumenDto> compras,
        CompraDto? compraEditar)
    {
        var items = compras
            .Select(x => new CompraResumenItemViewModel
            {
                IdCompra = x.IdCompra,
                NombreProveedor = x.NombreProveedor,
                EscenarioOperacion = x.EscenarioOperacion,
                FechaEmision = x.FechaEmision,
                FechaContabilizacion = x.FechaContabilizacion,
                Documento = $"{x.TipoComprobante} {x.Serie}-{x.Numero}",
                TipoComprobante = x.TipoComprobante,
                CodigoMoneda = x.CodigoMoneda,
                ImporteTotal = x.ImporteTotal,
                Saldo = x.Saldo,
                IdAsiento = x.IdAsiento,
                NumeroAsiento = x.NumeroAsiento,
                FechaValidacionCpe = x.FechaValidacionCpe,
                EstadoValidacionCpe = x.EstadoValidacionCpe,
                MensajeValidacionCpe = x.MensajeValidacionCpe,
                Estado = x.Estado,
                Situacion = x.Situacion
            })
            .ToList();

        var proveedorSeleccionado = proveedores.FirstOrDefault(x => x.IdProveedor == (compraEditar?.IdProveedor ?? proveedores.FirstOrDefault()?.IdProveedor));
        var totalImportePeriodo = items.Sum(x => x.ImporteTotal);
        var totalImporteSolesPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "PEN", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.ImporteTotal);
        var totalImporteDolaresPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "USD", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.ImporteTotal);

        return new CompraIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TipoComprobanteFiltro = tipoComprobanteFiltro ?? string.Empty,
            TotalCompras = items.Count,
            TotalImportePeriodo = totalImportePeriodo,
            TotalImporteSolesPeriodo = totalImporteSolesPeriodo,
            TotalImporteDolaresPeriodo = totalImporteDolaresPeriodo,
            PorcentajeRetencionRenta4ta = porcentajeRetencionRenta4ta,
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            Proveedores = proveedores.ToList(),
            ConfiguracionesCompra = configuraciones.ToList(),
            TiposDocumentoIdentidad = tiposDocumentoIdentidad.ToList(),
            Monedas = monedas.ToList(),
            TiposComprobante = tiposComprobante.ToList(),
            CuentasMovimiento = cuentasMovimiento.ToList(),
            TiposAfectacionIgv = tiposAfectacionIgv.ToList(),
            DetraccionesSunat = detraccionesSunat.ToList(),
            TiposPercepcion = tiposPercepcion.ToList(),
            Compras = items,
            ProveedorSeleccionadoTipoDocumento = proveedorSeleccionado?.TipoDocumento ?? compraEditar?.TipoDocumentoProveedor ?? string.Empty,
            ProveedorSeleccionadoNumeroDocumento = proveedorSeleccionado?.NumeroDocumento ?? compraEditar?.NumeroDocumentoProveedor ?? string.Empty,
            ProveedorSeleccionadoNombreLegal = proveedorSeleccionado?.NombreCompleto ?? string.Empty,
            ProveedorSeleccionadoTexto = proveedorSeleccionado is null ? string.Empty : $"{proveedorSeleccionado.NombreCompleto} ({proveedorSeleccionado.NumeroDocumento})",
            ProveedorSeleccionadoContacto = proveedorSeleccionado?.Contacto ?? proveedorSeleccionado?.Telefono ?? string.Empty,
            ProveedorSeleccionadoCorreo = proveedorSeleccionado?.CorreoElectronico ?? string.Empty,
            Formulario = compraEditar is null
                ? new CompraFormViewModel
                {
                    FechaEmision = ParsePeriodo(periodo),
                    FechaContabilizacion = ParsePeriodo(periodo),
                    IdMoneda = monedas.OrderByDescending(x => x.EsMonedaBase).FirstOrDefault()?.IdMoneda,
                    IdProveedor = proveedores.FirstOrDefault()?.IdProveedor,
                    IdConfiguracionContabilizacion = configuraciones.FirstOrDefault()?.IdConfiguracionContabilizacion,
                    TipoComprobante = tiposComprobante.FirstOrDefault()?.CodigoTipoComprobante ?? "01",
                    ExoneracionRenta4ta = false,
                    PorcentajeRetencion = porcentajeRetencionRenta4ta,
                    Retencion = 0,
                    TieneDetraccion = false,
                    PorcentajeDetraccion = 0,
                    ImporteDetraccion = 0,
                    TienePercepcion = false,
                    PorcentajePercepcion = 0,
                    BasePercepcion = 0,
                    ImportePercepcion = 0,
                    SaldoPago = 0,
                    Detalles =
                    [
                        new()
                        {
                            Item = 1,
                            IdTipoAfectacionIGV = tiposAfectacionIgv.FirstOrDefault(x => x.CodigoSunat == CodigoAfectacionGravadoOnerosa)?.IdTipoAfectacionIGV ?? 1
                        }
                    ]
                }
                : new CompraFormViewModel
                {
                    IdCompra = compraEditar.IdCompra,
                    IdProveedor = compraEditar.IdProveedor,
                    IdConfiguracionContabilizacion = compraEditar.IdConfiguracionContabilizacion,
                    FechaEmision = compraEditar.FechaEmision,
                    FechaContabilizacion = compraEditar.FechaContabilizacion,
                    TipoComprobante = compraEditar.TipoComprobante,
                    Serie = compraEditar.Serie,
                    Numero = compraEditar.Numero,
                    IdMoneda = compraEditar.IdMoneda,
                    TipoCambio = compraEditar.TipoCambio,
                    BaseImponible = compraEditar.BaseImponible,
                    TotalExonerado = compraEditar.TotalExonerado,
                    TotalInafecto = compraEditar.TotalInafecto,
                    Icbper = compraEditar.Icbper,
                    Igv = compraEditar.Igv,
                    Isc = compraEditar.Isc,
                    OtrosTributos = compraEditar.OtrosTributos,
                    Redondeo = compraEditar.Redondeo,
                    ImporteTotal = compraEditar.ImporteTotal,
                    ExoneracionRenta4ta = compraEditar.ExoneracionRenta4ta,
                    PorcentajeRetencion = compraEditar.PorcentajeRetencion,
                    Retencion = compraEditar.Retencion,
                    TieneDetraccion = compraEditar.TieneDetraccion,
                    IdDetraccionSunat = compraEditar.IdDetraccionSunat,
                    DetraccionTexto = string.IsNullOrWhiteSpace(compraEditar.CodigoDetraccionSunat)
                        ? string.Empty
                        : $"{compraEditar.CodigoDetraccionSunat} - {compraEditar.DescripcionDetraccionSunat}",
                    PorcentajeDetraccion = compraEditar.PorcentajeDetraccion,
                    ImporteDetraccion = compraEditar.ImporteDetraccion,
                    TienePercepcion = compraEditar.TienePercepcion,
                    IdTipoPercepcion = compraEditar.IdTipoPercepcion,
                    TipoPercepcionTexto = string.IsNullOrWhiteSpace(compraEditar.CodigoPercepcion)
                        ? string.Empty
                        : $"{compraEditar.CodigoPercepcion} - {compraEditar.DescripcionPercepcion}",
                    PorcentajePercepcion = compraEditar.PorcentajePercepcion,
                    BasePercepcion = compraEditar.BasePercepcion,
                    ImportePercepcion = compraEditar.ImportePercepcion,
                    SaldoPago = compraEditar.Saldo,
                    Observacion = compraEditar.Observacion,
                    Detalles = compraEditar.Detalles
                        .OrderBy(x => x.Item)
                        .Select(x => new CompraDetalleFormViewModel
                        {
                            Item = x.Item,
                            IdPlanCuenta = x.IdPlanCuenta,
                            CuentaTexto = string.IsNullOrWhiteSpace(x.CodigoCuenta) ? string.Empty : $"{x.CodigoCuenta} - {x.NombreCuenta}",
                            IdTipoAfectacionIGV = x.IdTipoAfectacionIGV,
                            Descripcion = x.Descripcion,
                            Cantidad = x.Cantidad,
                            ValorUnitario = x.ValorUnitario,
                            ImporteBruto = x.ImporteBruto
                        })
                        .ToList()
                }
        };
    }

    private async Task<decimal> ObtenerPorcentajeRetencionRenta4taAsync(int empresaId, CancellationToken cancellationToken)
    {
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, CodigoParametroPorcentajeRetencion4ta, 1, 20, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x =>
            x.Activo &&
            string.Equals(x.CodigoParametro, CodigoParametroPorcentajeRetencion4ta, StringComparison.OrdinalIgnoreCase));

        if (parametro is null
            || string.IsNullOrWhiteSpace(parametro.ValorParametro)
            || !decimal.TryParse(parametro.ValorParametro.Trim(), out var porcentaje))
        {
            return 0m;
        }

        return decimal.Round(porcentaje, 4);
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

    private static DateOnly ParsePeriodo(string periodo)
    {
        if (periodo.Length == 6
            && int.TryParse(periodo[..4], out var year)
            && int.TryParse(periodo[4..], out var month)
            && month is >= 1 and <= 12)
        {
            return new DateOnly(year, month, 1);
        }

        return DateOnly.FromDateTime(DateTime.Today);
    }

    private static (short anio, byte mes) DescomponerPeriodo(string periodo)
    {
        return (short.Parse(periodo[..4]), byte.Parse(periodo[4..]));
    }

    private static string NormalizarSerieDocumento(string? serie)
    {
        var serieNormalizada = new string((serie ?? string.Empty)
            .Trim()
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());

        if (string.IsNullOrEmpty(serieNormalizada))
        {
            return string.Empty;
        }

        return serieNormalizada.Length > 10 ? serieNormalizada[..10] : serieNormalizada;
    }

    private static string NormalizarNumeroDocumento(string? numero)
    {
        var normalizado = new string((numero ?? string.Empty)
            .Trim()
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());
        if (string.IsNullOrEmpty(normalizado))
        {
            return string.Empty;
        }

        return normalizado.Length > 20 ? normalizado[..20] : normalizado;
    }

    private static bool EsAfectacionGravada(string? codigoSunat)
        => !string.IsNullOrWhiteSpace(codigoSunat) && codigoSunat.Trim().StartsWith('1');

    private static bool EsAfectacionExonerada(string? codigoSunat)
        => !string.IsNullOrWhiteSpace(codigoSunat) && codigoSunat.Trim().StartsWith('2');

    private static bool EsAfectacionInafecta(string? codigoSunat)
        => !string.IsNullOrWhiteSpace(codigoSunat) && codigoSunat.Trim().StartsWith('3');

    private static string DescribirEstadoComprobante(string? codigo)
    {
        return codigo switch
        {
            "0" => "0 - NO EXISTE",
            "1" => "1 - ACEPTADO",
            "2" => "2 - ANULADO",
            "3" => "3 - AUTORIZADO",
            "4" => "4 - NO AUTORIZADO",
            _ => string.IsNullOrWhiteSpace(codigo) ? "SIN ESTADO" : codigo.Trim()
        };
    }

    private static string ConstruirMensajeValidacionCpe(MigoCpeResultDto resultado)
    {
        if (!string.IsNullOrWhiteSpace(resultado.Observaciones))
        {
            return resultado.Observaciones!;
        }

        var partes = new List<string>();
        if (!string.IsNullOrWhiteSpace(resultado.EstadoContribuyente))
        {
            partes.Add($"Estado contribuyente: {resultado.EstadoContribuyente}");
        }
        if (!string.IsNullOrWhiteSpace(resultado.CondicionDomicilio))
        {
            partes.Add($"Condicion domicilio: {resultado.CondicionDomicilio}");
        }

        return partes.Count == 0
            ? "Consulta realizada correctamente en SUNAT a traves de Migo."
            : string.Join(" | ", partes);
    }
}
