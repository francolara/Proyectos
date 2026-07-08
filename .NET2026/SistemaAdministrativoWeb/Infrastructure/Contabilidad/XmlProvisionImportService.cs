using System.Globalization;
using System.Xml.Linq;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class XmlProvisionImportService(
    ICompraRepository compraRepository,
    IVentaRepository ventaRepository,
    IProveedorRepository proveedorRepository,
    IClienteRepository clienteRepository,
    IPersonaRepository personaRepository,
    IConfiguracionContabilizacionRepository configuracionContabilizacionRepository,
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository,
    ITipoAfectacionIgvRepository tipoAfectacionIgvRepository,
    ITipoCambioRepository tipoCambioRepository,
    ITipoCambioSyncService tipoCambioSyncService,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    IParametroEmpresaRepository parametroEmpresaRepository,
    IDetraccionSunatRepository detraccionSunatRepository,
    IPlanCuentaRepository planCuentaRepository,
    ITipoPercepcionRepository tipoPercepcionRepository) : IXmlProvisionImportService
{
    private const string CodigoMonedaUsd = "USD";
    private const string CodigoMonedaPen = "PEN";
    private const string CodigoReciboHonorarios = "02";
    private const string CodigoParametroPorcentajeRetencion4ta = "PORCRETEN4TA";
    private const string CodigoParametroCuentaCompraDefault = "CTACOMPRADEFAULT";
    private const string CodigoParametroCuentaVentaDefault = "CTAVENTADEFAULT";

    public async Task<ImportacionXmlResultadoDto> ImportarComprasAsync(int idEmpresa, IReadOnlyCollection<ImportacionXmlArchivoRequest> archivos, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        return await ImportarAsync(idEmpresa, archivos, usuarioRegistro, true, cancellationToken);
    }

    public async Task<ImportacionXmlResultadoDto> ImportarVentasAsync(int idEmpresa, IReadOnlyCollection<ImportacionXmlArchivoRequest> archivos, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        return await ImportarAsync(idEmpresa, archivos, usuarioRegistro, false, cancellationToken);
    }

    private async Task<ImportacionXmlResultadoDto> ImportarAsync(int idEmpresa, IReadOnlyCollection<ImportacionXmlArchivoRequest> archivos, string? usuarioRegistro, bool esCompra, CancellationToken cancellationToken)
    {
        var resultado = new ImportacionXmlResultadoDto();
        if (archivos.Count == 0)
        {
            resultado.Items.Add(new ImportacionXmlResultadoItemDto
            {
                NombreArchivo = string.Empty,
                Importado = false,
                Mensaje = "Seleccione al menos un archivo XML."
            });
            return resultado;
        }

        var configuraciones = (await configuracionContabilizacionRepository.ListarPorEmpresaAsync(idEmpresa, cancellationToken))
            .Where(x => x.Activo && x.GeneraAsientoAutomatico && x.ModuloOperacion == (esCompra ? "COM" : "VEN"))
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var configuracionDefault = configuraciones.FirstOrDefault();
        if (configuracionDefault is null)
        {
            throw new InvalidOperationException($"No existe una configuracion contable activa para {(esCompra ? "compras" : "ventas")}.");
        }

        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .ToDictionary(x => x.CodigoMoneda.Trim().ToUpperInvariant(), StringComparer.OrdinalIgnoreCase);
        var tiposAfectacion = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .ToDictionary(x => x.CodigoSunat.Trim(), x => x, StringComparer.OrdinalIgnoreCase);
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(esCompra, !esCompra, cancellationToken))
            .ToDictionary(x => x.CodigoTipoComprobante.Trim(), x => x, StringComparer.OrdinalIgnoreCase);
        var detracciones = esCompra
            ? (await detraccionSunatRepository.ListarActivasAsync(cancellationToken))
                .ToDictionary(x => x.CodigoSunat.Trim(), x => x, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, DetraccionSunatDto>(StringComparer.OrdinalIgnoreCase);
        var percepciones = esCompra
            ? (await tipoPercepcionRepository.ListarActivasAsync(cancellationToken))
                .ToDictionary(x => x.Codigo.Trim(), x => x, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, TipoPercepcionDto>(StringComparer.OrdinalIgnoreCase);
        var contextoSuscripcion = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(idEmpresa, cancellationToken);

        if (contextoSuscripcion is null)
        {
            throw new InvalidOperationException("No se encontro la cuenta administradora de la empresa activa.");
        }

        var porcentajeRetencion4ta = esCompra
            ? await ObtenerPorcentajeRetencionRenta4taAsync(idEmpresa, cancellationToken)
            : 0m;
        var idPlanCuentaDefault = await ObtenerCuentaDefaultImportacionAsync(
            idEmpresa,
            esCompra ? CodigoParametroCuentaCompraDefault : CodigoParametroCuentaVentaDefault,
            esCompra ? "compras" : "ventas",
            cancellationToken);

        foreach (var archivo in archivos)
        {
            ComprobanteXmlImportado? comprobante = null;
            try
            {
                comprobante = ParsearComprobante(archivo, esCompra, tiposAfectacion);

                if (!tiposComprobante.ContainsKey(comprobante.TipoComprobante))
                {
                    throw new InvalidOperationException($"El tipo de comprobante {comprobante.TipoComprobante} no esta habilitado para {(esCompra ? "compras" : "ventas")}.");
                }

                if (!monedas.TryGetValue(comprobante.CodigoMoneda, out var moneda))
                {
                    throw new InvalidOperationException($"La moneda {comprobante.CodigoMoneda} no esta registrada en el sistema.");
                }

                var tipoCambio = await ResolverTipoCambioAsync(contextoSuscripcion.IdCuentaAdministradora, comprobante.FechaEmision, cancellationToken);
                var tercero = esCompra
                    ? await ResolverProveedorAsync(idEmpresa, comprobante, usuarioRegistro, cancellationToken)
                    : await ResolverClienteAsync(idEmpresa, comprobante, usuarioRegistro, cancellationToken);

                if (esCompra)
                {
                    var detraccion = comprobante.CodigoDetraccion is not null && detracciones.TryGetValue(comprobante.CodigoDetraccion, out var detraccionEncontrada)
                        ? detraccionEncontrada
                        : null;
                    var idDetraccionSunat = detraccion is not null
                        ? detraccion.IdDetraccionSunat
                        : (int?)null;
                    var porcentajeDetraccion = detraccion?.Porcentaje ?? 0m;
                    var importeDetraccion = comprobante.TieneDetraccion && porcentajeDetraccion > 0
                        ? decimal.Round(comprobante.ImporteTotal * (porcentajeDetraccion / 100m), 2)
                        : 0m;
                    var percepcion = comprobante.CodigoPercepcion is not null && percepciones.TryGetValue(comprobante.CodigoPercepcion, out var percepcionEncontrada)
                        ? percepcionEncontrada
                        : null;
                    var idTipoPercepcion = percepcion is not null
                        ? percepcion.IdTipoPercepcion
                        : (int?)null;
                    var porcentajePercepcion = percepcion?.Porcentaje ?? 0m;
                    var basePercepcion = comprobante.TienePercepcion ? comprobante.ImporteTotal : 0m;
                    var importePercepcion = comprobante.TienePercepcion && porcentajePercepcion > 0
                        ? decimal.Round(basePercepcion * (porcentajePercepcion / 100m), 2)
                        : 0m;
                    var esReciboHonorarios = string.Equals(comprobante.TipoComprobante, CodigoReciboHonorarios, StringComparison.OrdinalIgnoreCase);
                    var retencion = esReciboHonorarios
                        ? decimal.Round(comprobante.BaseImponible * (porcentajeRetencion4ta / 100m), 2)
                        : 0m;
                    var importeTotal = esReciboHonorarios
                        ? decimal.Round(comprobante.BaseImponible - retencion, 2)
                        : comprobante.ImporteTotal;

                    var guardado = await compraRepository.ImportarXmlAsync(new ImportarCompraXmlRequest
                    {
                        IdEmpresa = idEmpresa,
                        IdProveedor = tercero.Id,
                        IdConfiguracionContabilizacion = configuracionDefault.IdConfiguracionContabilizacion,
                        FechaEmision = comprobante.FechaEmision,
                        FechaContabilizacion = comprobante.FechaEmision,
                        TipoComprobante = comprobante.TipoComprobante,
                        Serie = comprobante.Serie,
                        Numero = comprobante.Numero,
                        IdMoneda = moneda.IdMoneda,
                        TipoCambio = tipoCambio,
                        BaseImponible = comprobante.BaseImponible,
                        TotalExonerado = comprobante.TotalExonerado,
                        TotalInafecto = comprobante.TotalInafecto,
                        Icbper = comprobante.Icbper,
                        Igv = esReciboHonorarios ? 0m : comprobante.Igv,
                        Isc = comprobante.Isc,
                        OtrosTributos = comprobante.OtrosTributos,
                        Redondeo = comprobante.Redondeo,
                        ImporteTotal = importeTotal,
                        ExoneracionRenta4ta = false,
                        PorcentajeRetencion = esReciboHonorarios ? porcentajeRetencion4ta : 0m,
                        Retencion = retencion,
                        TieneDetraccion = comprobante.TieneDetraccion && idDetraccionSunat.HasValue,
                        IdDetraccionSunat = idDetraccionSunat,
                        PorcentajeDetraccion = porcentajeDetraccion,
                        ImporteDetraccion = importeDetraccion,
                        TienePercepcion = comprobante.TienePercepcion && idTipoPercepcion.HasValue,
                        IdTipoPercepcion = idTipoPercepcion,
                        PorcentajePercepcion = porcentajePercepcion,
                        BasePercepcion = basePercepcion,
                        ImportePercepcion = importePercepcion,
                        Observacion = $"Importado desde XML {archivo.NombreArchivo}",
                        UsuarioRegistro = usuarioRegistro,
                        Detalles = comprobante.Detalles
                            .Select(x => new ImportarCompraXmlDetalleRequest
                            {
                                Item = x.Item,
                                IdPlanCuenta = idPlanCuentaDefault,
                                IdTipoAfectacionIGV = x.IdTipoAfectacionIGV,
                                Descripcion = x.Descripcion,
                                Cantidad = x.Cantidad,
                                ValorUnitario = x.ValorUnitario,
                                ImporteBruto = x.ImporteBruto
                            })
                            .ToList()
                    }, cancellationToken);

                    resultado.Items.Add(new ImportacionXmlResultadoItemDto
                    {
                        NombreArchivo = archivo.NombreArchivo,
                        Importado = true,
                        Mensaje = $"Compra importada en estado {guardado.Estado}.",
                        TipoComprobante = comprobante.TipoComprobante,
                        Serie = comprobante.Serie,
                        Numero = comprobante.Numero,
                        NombreTercero = tercero.Nombre,
                        ImporteTotal = guardado.ImporteTotal,
                        IdRegistro = guardado.IdCompra
                    });
                }
                else
                {
                    var guardado = await ventaRepository.ImportarXmlAsync(new ImportarVentaXmlRequest
                    {
                        IdEmpresa = idEmpresa,
                        IdCliente = tercero.Id,
                        IdConfiguracionContabilizacion = configuracionDefault.IdConfiguracionContabilizacion,
                        FechaEmision = comprobante.FechaEmision,
                        FechaContabilizacion = comprobante.FechaEmision,
                        TipoComprobante = comprobante.TipoComprobante,
                        Serie = comprobante.Serie,
                        Numero = comprobante.Numero,
                        IdMoneda = moneda.IdMoneda,
                        TipoCambio = tipoCambio,
                        BaseImponible = comprobante.BaseImponible,
                        TotalExonerado = comprobante.TotalExonerado,
                        TotalInafecto = comprobante.TotalInafecto,
                        Icbper = comprobante.Icbper,
                        Igv = comprobante.Igv,
                        Isc = comprobante.Isc,
                        OtrosTributos = comprobante.OtrosTributos,
                        Redondeo = comprobante.Redondeo,
                        ImporteTotal = comprobante.ImporteTotal,
                        Observacion = $"Importado desde XML {archivo.NombreArchivo}",
                        UsuarioRegistro = usuarioRegistro,
                        Detalles = comprobante.Detalles
                            .Select(x => new ImportarVentaXmlDetalleRequest
                            {
                                Item = x.Item,
                                IdPlanCuenta = idPlanCuentaDefault,
                                IdTipoAfectacionIGV = x.IdTipoAfectacionIGV,
                                Descripcion = x.Descripcion,
                                Cantidad = x.Cantidad,
                                ValorUnitario = x.ValorUnitario,
                                ImporteBruto = x.ImporteBruto
                            })
                            .ToList()
                    }, cancellationToken);

                    resultado.Items.Add(new ImportacionXmlResultadoItemDto
                    {
                        NombreArchivo = archivo.NombreArchivo,
                        Importado = true,
                        Mensaje = $"Venta importada en estado {guardado.Estado}.",
                        TipoComprobante = comprobante.TipoComprobante,
                        Serie = comprobante.Serie,
                        Numero = comprobante.Numero,
                        NombreTercero = tercero.Nombre,
                        ImporteTotal = guardado.ImporteTotal,
                        IdRegistro = guardado.IdVenta
                    });
                }
            }
            catch (Exception ex)
            {
                resultado.Items.Add(new ImportacionXmlResultadoItemDto
                {
                    NombreArchivo = archivo.NombreArchivo,
                    Importado = false,
                    Mensaje = ex.Message,
                    TipoComprobante = comprobante?.TipoComprobante ?? string.Empty,
                    Serie = comprobante?.Serie ?? string.Empty,
                    Numero = comprobante?.Numero ?? string.Empty,
                    NombreTercero = comprobante?.NombreTercero ?? string.Empty,
                    ImporteTotal = ObtenerImporteTotalResultado(comprobante, esCompra, porcentajeRetencion4ta)
                });
            }
        }

        return resultado;
    }

    private static decimal ObtenerImporteTotalResultado(ComprobanteXmlImportado? comprobante, bool esCompra, decimal porcentajeRetencion4ta)
    {
        if (comprobante is null)
        {
            return 0m;
        }

        var esReciboHonorariosCompra = esCompra
            && string.Equals(comprobante.TipoComprobante, CodigoReciboHonorarios, StringComparison.OrdinalIgnoreCase);

        if (!esReciboHonorariosCompra)
        {
            return comprobante.ImporteTotal;
        }

        var retencion = decimal.Round(comprobante.BaseImponible * (porcentajeRetencion4ta / 100m), 2);
        return decimal.Round(comprobante.BaseImponible - retencion, 2);
    }

    private async Task<decimal> ResolverTipoCambioAsync(int idCuentaAdministradora, DateOnly fecha, CancellationToken cancellationToken)
    {
        var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(idCuentaAdministradora, fecha, CodigoMonedaUsd, cancellationToken)
            ?? await tipoCambioSyncService.SincronizarFechaAsync(idCuentaAdministradora, fecha, CodigoMonedaUsd, null, cancellationToken);

        if (tipoCambio is null || tipoCambio.Venta <= 0)
        {
            throw new InvalidOperationException($"No existe tipo de cambio USD para la fecha {fecha:dd/MM/yyyy}.");
        }

        return decimal.Round(tipoCambio.Venta, 3);
    }

    private async Task<(int Id, string Nombre)> ResolverProveedorAsync(int idEmpresa, ComprobanteXmlImportado comprobante, string? usuarioRegistro, CancellationToken cancellationToken)
    {
        var proveedores = await proveedorRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
        var existente = proveedores.FirstOrDefault(x =>
            string.Equals(x.TipoDocumento, comprobante.TipoDocumentoTercero, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.NumeroDocumento, comprobante.NumeroDocumentoTercero, StringComparison.OrdinalIgnoreCase));

        if (existente is not null)
        {
            return (existente.IdProveedor, existente.NombreCompleto);
        }

        await personaRepository.GuardarAsync(new GuardarPersonaRequest
        {
            IdEmpresa = idEmpresa,
            TipoPersona = DeterminarTipoPersona(comprobante, true),
            TipoDocumento = comprobante.TipoDocumentoTercero,
            NumeroDocumento = comprobante.NumeroDocumentoTercero,
            ApellidoPaterno = null,
            ApellidoMaterno = null,
            Nombres = DeterminarTipoPersona(comprobante, true) == "N" ? comprobante.NombreTercero : null,
            RazonSocial = DeterminarTipoPersona(comprobante, true) == "J" ? comprobante.NombreTercero : null,
            Direccion = comprobante.DireccionTercero,
            CodigoUbigeo = comprobante.UbigeoTercero,
            EsCliente = false,
            EsProveedor = true,
            Estado = true,
            UsuarioRegistro = usuarioRegistro
        }, cancellationToken);

        proveedores = await proveedorRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
        existente = proveedores.FirstOrDefault(x =>
            string.Equals(x.TipoDocumento, comprobante.TipoDocumentoTercero, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.NumeroDocumento, comprobante.NumeroDocumentoTercero, StringComparison.OrdinalIgnoreCase));

        if (existente is null)
        {
            throw new InvalidOperationException($"No se pudo crear o recuperar el proveedor {comprobante.NombreTercero}.");
        }

        return (existente.IdProveedor, existente.NombreCompleto);
    }

    private async Task<(int Id, string Nombre)> ResolverClienteAsync(int idEmpresa, ComprobanteXmlImportado comprobante, string? usuarioRegistro, CancellationToken cancellationToken)
    {
        var clientes = await clienteRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
        var existente = clientes.FirstOrDefault(x =>
            string.Equals(x.TipoDocumento, comprobante.TipoDocumentoTercero, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.NumeroDocumento, comprobante.NumeroDocumentoTercero, StringComparison.OrdinalIgnoreCase));

        if (existente is not null)
        {
            return (existente.IdCliente, existente.NombreCompleto);
        }

        await personaRepository.GuardarAsync(new GuardarPersonaRequest
        {
            IdEmpresa = idEmpresa,
            TipoPersona = DeterminarTipoPersona(comprobante, false),
            TipoDocumento = comprobante.TipoDocumentoTercero,
            NumeroDocumento = comprobante.NumeroDocumentoTercero,
            ApellidoPaterno = null,
            ApellidoMaterno = null,
            Nombres = DeterminarTipoPersona(comprobante, false) == "N" ? comprobante.NombreTercero : null,
            RazonSocial = DeterminarTipoPersona(comprobante, false) == "J" ? comprobante.NombreTercero : null,
            Direccion = comprobante.DireccionTercero,
            CodigoUbigeo = comprobante.UbigeoTercero,
            EsCliente = true,
            EsProveedor = false,
            Estado = true,
            UsuarioRegistro = usuarioRegistro
        }, cancellationToken);

        clientes = await clienteRepository.ListarActivosPorEmpresaAsync(idEmpresa, cancellationToken);
        existente = clientes.FirstOrDefault(x =>
            string.Equals(x.TipoDocumento, comprobante.TipoDocumentoTercero, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.NumeroDocumento, comprobante.NumeroDocumentoTercero, StringComparison.OrdinalIgnoreCase));

        if (existente is null)
        {
            throw new InvalidOperationException($"No se pudo crear o recuperar el cliente {comprobante.NombreTercero}.");
        }

        return (existente.IdCliente, existente.NombreCompleto);
    }

    private static string DeterminarTipoPersona(ComprobanteXmlImportado comprobante, bool esCompra)
    {
        if (string.Equals(comprobante.TipoComprobante, CodigoReciboHonorarios, StringComparison.OrdinalIgnoreCase) && esCompra)
        {
            return "N";
        }

        return string.Equals(comprobante.TipoDocumentoTercero, "6", StringComparison.OrdinalIgnoreCase) ? "J" : "N";
    }

    private async Task<decimal> ObtenerPorcentajeRetencionRenta4taAsync(int idEmpresa, CancellationToken cancellationToken)
    {
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, null, CodigoParametroPorcentajeRetencion4ta, 1, 20, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x =>
            x.Activo &&
            string.Equals(x.CodigoParametro, CodigoParametroPorcentajeRetencion4ta, StringComparison.OrdinalIgnoreCase));

        if (parametro is null
            || string.IsNullOrWhiteSpace(parametro.ValorParametro)
            || !decimal.TryParse(parametro.ValorParametro.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var porcentaje)
                && !decimal.TryParse(parametro.ValorParametro.Trim(), out porcentaje))
        {
            return 0m;
        }

        return decimal.Round(porcentaje, 4);
    }

    private async Task<int> ObtenerCuentaDefaultImportacionAsync(int idEmpresa, string codigoParametro, string modulo, CancellationToken cancellationToken)
    {
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, null, codigoParametro, 1, 20, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x =>
            x.Activo &&
            string.Equals(x.CodigoParametro, codigoParametro, StringComparison.OrdinalIgnoreCase));

        if (parametro is null
            || string.IsNullOrWhiteSpace(parametro.ValorParametro))
        {
            throw new InvalidOperationException($"No existe una cuenta contable default valida configurada en el parametro {codigoParametro} para la importacion de {modulo}.");
        }

        var valorParametro = parametro.ValorParametro.Trim();
        var cuentas = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, valorParametro, null, 1, 20, false, false, cancellationToken);
        var cuenta = cuentas.Items.FirstOrDefault(x => string.Equals(x.CodigoCuenta, valorParametro, StringComparison.OrdinalIgnoreCase));
        if (cuenta is not null && cuenta.IdPlanCuenta > 0)
        {
            return cuenta.IdPlanCuenta;
        }

        if (int.TryParse(valorParametro, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idPlanCuenta) && idPlanCuenta > 0)
        {
            var cuentasPorId = await planCuentaRepository.ListarPorEmpresaAsync(idEmpresa, false, cancellationToken);
            var cuentaPorId = cuentasPorId.FirstOrDefault(x => x.IdPlanCuenta == idPlanCuenta);
            if (cuentaPorId is not null)
            {
                return cuentaPorId.IdPlanCuenta;
            }
        }

        throw new InvalidOperationException($"La cuenta configurada en el parametro {codigoParametro} no existe en el plan de cuentas de la empresa para la importacion de {modulo}.");
    }

    private static ComprobanteXmlImportado ParsearComprobante(ImportacionXmlArchivoRequest archivo, bool esCompra, IReadOnlyDictionary<string, TipoAfectacionIgvDto> tiposAfectacion)
    {
        using var stream = new MemoryStream(archivo.Contenido, writable: false);
        var document = XDocument.Load(stream, LoadOptions.None);
        var root = document.Root ?? throw new InvalidOperationException($"El archivo {archivo.NombreArchivo} no contiene un XML valido.");

        var tipoComprobante = ObtenerTipoComprobante(root);
        ValidarTipoPermitido(tipoComprobante, esCompra);

        var idDocumento = ObtenerValorPrimerDescendiente(root, "ID");
        var (serie, numero) = SepararDocumento(idDocumento);
        var fechaEmisionTexto = ObtenerValorPrimerDescendiente(root, "IssueDate");
        if (!DateOnly.TryParse(fechaEmisionTexto, out var fechaEmision))
        {
            throw new InvalidOperationException($"El archivo {archivo.NombreArchivo} no contiene una fecha de emision valida.");
        }

        var tercero = esCompra
            ? ObtenerTercero(root, "AccountingSupplierParty")
            : ObtenerTercero(root, "AccountingCustomerParty");

        var detalle = ObtenerDetalles(root, tiposAfectacion);
        if (detalle.Count == 0)
        {
            throw new InvalidOperationException($"El archivo {archivo.NombreArchivo} no contiene detalle de lineas.");
        }

        var totales = ObtenerTotales(root, detalle, tipoComprobante);
        return new ComprobanteXmlImportado
        {
            TipoComprobante = tipoComprobante,
            Serie = NormalizarSerie(serie),
            Numero = NormalizarNumero(numero),
            FechaEmision = fechaEmision,
            CodigoMoneda = (ObtenerValorPrimerDescendiente(root, "DocumentCurrencyCode") ?? CodigoMonedaPen).Trim().ToUpperInvariant(),
            TipoDocumentoTercero = tercero.TipoDocumento,
            NumeroDocumentoTercero = tercero.NumeroDocumento,
            NombreTercero = tercero.Nombre,
            DireccionTercero = tercero.Direccion,
            UbigeoTercero = tercero.Ubigeo,
            BaseImponible = totales.BaseImponible,
            TotalExonerado = totales.TotalExonerado,
            TotalInafecto = totales.TotalInafecto,
            Icbper = totales.Icbper,
            Igv = totales.Igv,
            Isc = totales.Isc,
            OtrosTributos = totales.OtrosTributos,
            Redondeo = totales.Redondeo,
            ImporteTotal = totales.ImporteTotal,
            TieneDetraccion = false,
            TienePercepcion = false,
            Detalles = detalle
        };
    }

    private static void ValidarTipoPermitido(string tipoComprobante, bool esCompra)
    {
        var permitidos = esCompra
            ? new[] { "01", "03", "07", "08", "02" }
            : new[] { "01", "03", "07", "08" };

        if (!permitidos.Contains(tipoComprobante, StringComparer.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"El tipo de comprobante {tipoComprobante} no esta permitido para {(esCompra ? "compras" : "ventas")}.");
        }
    }

    private static string ObtenerTipoComprobante(XElement root)
    {
        var rootName = root.Name.LocalName;
        if (string.Equals(rootName, "Invoice", StringComparison.OrdinalIgnoreCase))
        {
            return (ObtenerValorPrimerDescendiente(root, "InvoiceTypeCode") ?? "01").Trim().ToUpperInvariant();
        }

        if (string.Equals(rootName, "CreditNote", StringComparison.OrdinalIgnoreCase))
        {
            return (ObtenerValorPrimerDescendiente(root, "CreditNoteTypeCode") ?? "07").Trim().ToUpperInvariant();
        }

        if (string.Equals(rootName, "DebitNote", StringComparison.OrdinalIgnoreCase))
        {
            return (ObtenerValorPrimerDescendiente(root, "DebitNoteTypeCode") ?? "08").Trim().ToUpperInvariant();
        }

        throw new InvalidOperationException($"El documento XML {rootName} no esta soportado.");
    }

    private static (string Serie, string Numero) SepararDocumento(string? documento)
    {
        var valor = (documento ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(valor))
        {
            throw new InvalidOperationException("El XML no contiene la serie y numero del comprobante.");
        }

        var partes = valor.Split('-', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (partes.Length != 2)
        {
            throw new InvalidOperationException($"El identificador del comprobante {valor} no tiene el formato serie-numero.");
        }

        return (partes[0], partes[1]);
    }

    private static string NormalizarSerie(string serie)
    {
        var normalizada = new string((serie ?? string.Empty)
            .Trim()
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());

        return string.IsNullOrWhiteSpace(normalizada)
            ? string.Empty
            : (normalizada.Length > 10 ? normalizada[..10] : normalizada);
    }

    private static string NormalizarNumero(string numero)
    {
        var normalizado = new string((numero ?? string.Empty).Where(char.IsLetterOrDigit).ToArray());
        return string.IsNullOrWhiteSpace(normalizado)
            ? string.Empty
            : (normalizado.Length > 20 ? normalizado[..20] : normalizado);
    }

    private static (string TipoDocumento, string NumeroDocumento, string Nombre, string? Direccion, string? Ubigeo) ObtenerTercero(XElement root, string nodoParty)
    {
        var party = root.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, nodoParty, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El XML no contiene el bloque {nodoParty}.");
        var idNode = party.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "ID", StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El XML no contiene el documento del tercero en {nodoParty}.");
        var tipoDocumento = (idNode.Attribute("schemeID")?.Value ?? "6").Trim();
        var numeroDocumento = (idNode.Value ?? string.Empty).Trim();
        var nombre = (party.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "RegistrationName", StringComparison.OrdinalIgnoreCase))?.Value
            ?? party.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value
            ?? string.Empty).Trim();
        var direccion = party.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "StreetName", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();
        var ubigeo = party.Descendants().FirstOrDefault(x =>
            string.Equals(x.Name.LocalName, "ID", StringComparison.OrdinalIgnoreCase)
            && string.Equals(x.Parent?.Name.LocalName, "Address", StringComparison.OrdinalIgnoreCase))?.Value?.Trim();

        if (string.IsNullOrWhiteSpace(numeroDocumento) || string.IsNullOrWhiteSpace(nombre))
        {
            throw new InvalidOperationException($"El XML no contiene datos completos del tercero en {nodoParty}.");
        }

        return (tipoDocumento, numeroDocumento, nombre, direccion, ubigeo);
    }

    private static List<ComprobanteXmlDetalleImportado> ObtenerDetalles(XElement root, IReadOnlyDictionary<string, TipoAfectacionIgvDto> tiposAfectacion)
    {
        var nombresLinea = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "InvoiceLine",
            "CreditNoteLine",
            "DebitNoteLine"
        };

        var detalle = new List<ComprobanteXmlDetalleImportado>();
        short item = 1;
        foreach (var linea in root.Descendants().Where(x => nombresLinea.Contains(x.Name.LocalName)))
        {
            var cantidad = ParseDecimal(ObtenerValorPrimerDescendiente(linea, "InvoicedQuantity")
                ?? ObtenerValorPrimerDescendiente(linea, "CreditedQuantity")
                ?? ObtenerValorPrimerDescendiente(linea, "DebitedQuantity")
                ?? "1");
            if (cantidad <= 0)
            {
                cantidad = 1;
            }

            var descripcion = (linea.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Description", StringComparison.OrdinalIgnoreCase))?.Value
                ?? linea.Descendants().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value
                ?? $"Item {item}").Trim();
            var importeBruto = ParseDecimal(ObtenerValorPrimerDescendiente(linea, "LineExtensionAmount") ?? "0");
            var valorUnitario = ParseDecimal(ObtenerValorPrimerDescendiente(linea, "PriceAmount") ?? "0");
            if (valorUnitario <= 0 && cantidad > 0)
            {
                valorUnitario = decimal.Round(importeBruto / cantidad, 6);
            }

            var codigoAfectacion = (ObtenerValorPrimerDescendiente(linea, "TaxExemptionReasonCode")
                ?? InferirCodigoAfectacionPorLinea(linea)).Trim().ToUpperInvariant();
            if (!tiposAfectacion.TryGetValue(codigoAfectacion, out var afectacion))
            {
                afectacion = tiposAfectacion.Values.FirstOrDefault(x => string.Equals(x.CodigoSunat, "10", StringComparison.OrdinalIgnoreCase))
                    ?? tiposAfectacion.Values.First();
            }

            detalle.Add(new ComprobanteXmlDetalleImportado
            {
                Item = item++,
                IdTipoAfectacionIGV = afectacion.IdTipoAfectacionIGV,
                CodigoAfectacionIGV = afectacion.CodigoSunat,
                Descripcion = descripcion,
                Cantidad = decimal.Round(cantidad, 4),
                ValorUnitario = decimal.Round(valorUnitario, 6),
                ImporteBruto = decimal.Round(importeBruto, 2)
            });
        }

        return detalle;
    }

    private static string InferirCodigoAfectacionPorLinea(XElement linea)
    {
        var montoIgv = linea.Descendants()
            .Where(x => string.Equals(x.Name.LocalName, "TaxAmount", StringComparison.OrdinalIgnoreCase))
            .Select(x => ParseDecimal(x.Value))
            .FirstOrDefault();

        return montoIgv > 0 ? "10" : "20";
    }

    private static TotalesXmlImportados ObtenerTotales(XElement root, IReadOnlyCollection<ComprobanteXmlDetalleImportado> detalles, string tipoComprobante)
    {
        var payableAmount = ParseDecimal(ObtenerValorPrimerDescendiente(root, "PayableAmount") ?? "0");
        var taxInclusiveAmount = ParseDecimal(ObtenerValorPrimerDescendiente(root, "TaxInclusiveAmount") ?? "0");
        var igv = 0m;
        var icbper = 0m;
        var isc = 0m;
        var otrosTributos = 0m;

        foreach (var taxSubtotal in root.Descendants().Where(x => string.Equals(x.Name.LocalName, "TaxSubtotal", StringComparison.OrdinalIgnoreCase)))
        {
            var codigoTributo = ObtenerValorPrimerDescendiente(taxSubtotal, "ID")?.Trim();
            var monto = ParseDecimal(ObtenerValorPrimerDescendiente(taxSubtotal, "TaxAmount") ?? "0");
            switch (codigoTributo)
            {
                case "1000":
                    igv += monto;
                    break;
                case "2000":
                    isc += monto;
                    break;
                case "7152":
                    icbper += monto;
                    break;
                default:
                    if (monto > 0)
                    {
                        otrosTributos += monto;
                    }
                    break;
            }
        }

        var totalGravado = decimal.Round(detalles.Where(x => x.CodigoAfectacionIGV.StartsWith("1", StringComparison.OrdinalIgnoreCase)).Sum(x => x.ImporteBruto), 2);
        var totalExonerado = decimal.Round(detalles.Where(x => x.CodigoAfectacionIGV.StartsWith("2", StringComparison.OrdinalIgnoreCase)).Sum(x => x.ImporteBruto), 2);
        var totalInafecto = decimal.Round(detalles.Where(x => x.CodigoAfectacionIGV.StartsWith("3", StringComparison.OrdinalIgnoreCase)).Sum(x => x.ImporteBruto), 2);
        var baseImponible = decimal.Round(detalles.Sum(x => x.ImporteBruto), 2);
        var importeTotal = payableAmount > 0
            ? payableAmount
            : taxInclusiveAmount > 0
                ? taxInclusiveAmount
                : decimal.Round(baseImponible + totalExonerado + totalInafecto + igv + isc + icbper + otrosTributos, 2);

        return new TotalesXmlImportados
        {
            BaseImponible = baseImponible,
            TotalExonerado = totalExonerado,
            TotalInafecto = totalInafecto,
            Igv = decimal.Round(igv, 2),
            Icbper = decimal.Round(icbper, 2),
            Isc = decimal.Round(isc, 2),
            OtrosTributos = decimal.Round(otrosTributos, 2),
            Redondeo = 0m,
            ImporteTotal = decimal.Round(importeTotal, 2)
        };
    }

    private static string? ObtenerValorPrimerDescendiente(XElement elemento, string localName)
    {
        return elemento
            .Descendants()
            .FirstOrDefault(x => string.Equals(x.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase))
            ?.Value;
    }

    private static decimal ParseDecimal(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return 0m;
        }

        return decimal.TryParse(valor.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0m;
    }

    private sealed class ComprobanteXmlImportado
    {
        public string TipoComprobante { get; init; } = string.Empty;
        public string Serie { get; init; } = string.Empty;
        public string Numero { get; init; } = string.Empty;
        public DateOnly FechaEmision { get; init; }
        public string CodigoMoneda { get; init; } = CodigoMonedaPen;
        public string TipoDocumentoTercero { get; init; } = string.Empty;
        public string NumeroDocumentoTercero { get; init; } = string.Empty;
        public string NombreTercero { get; init; } = string.Empty;
        public string? DireccionTercero { get; init; }
        public string? UbigeoTercero { get; init; }
        public decimal BaseImponible { get; init; }
        public decimal TotalExonerado { get; init; }
        public decimal TotalInafecto { get; init; }
        public decimal Icbper { get; init; }
        public decimal Igv { get; init; }
        public decimal Isc { get; init; }
        public decimal OtrosTributos { get; init; }
        public decimal Redondeo { get; init; }
        public decimal ImporteTotal { get; init; }
        public bool TieneDetraccion { get; init; }
        public string? CodigoDetraccion { get; init; }
        public bool TienePercepcion { get; init; }
        public string? CodigoPercepcion { get; init; }
        public List<ComprobanteXmlDetalleImportado> Detalles { get; init; } = [];
    }

    private sealed class ComprobanteXmlDetalleImportado
    {
        public short Item { get; init; }
        public int IdTipoAfectacionIGV { get; init; }
        public string CodigoAfectacionIGV { get; init; } = string.Empty;
        public string Descripcion { get; init; } = string.Empty;
        public decimal Cantidad { get; init; }
        public decimal ValorUnitario { get; init; }
        public decimal ImporteBruto { get; init; }
    }

    private sealed class TotalesXmlImportados
    {
        public decimal BaseImponible { get; init; }
        public decimal TotalExonerado { get; init; }
        public decimal TotalInafecto { get; init; }
        public decimal Icbper { get; init; }
        public decimal Igv { get; init; }
        public decimal Isc { get; init; }
        public decimal OtrosTributos { get; init; }
        public decimal Redondeo { get; init; }
        public decimal ImporteTotal { get; init; }
    }
}
