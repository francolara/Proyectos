using System.Data;
using System.Globalization;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public sealed class ComprobanteElectronicoEmisionService(
    IConfiguration configuration,
    ISportCenterStoredProcedureService spService,
    IHttpClientFactory httpClientFactory,
    ILogger<ComprobanteElectronicoEmisionService> logger) : IComprobanteElectronicoEmisionService
{
    private readonly string _connectionString = configuration.GetConnectionString("DefaultConnection")
                                               ?? throw new InvalidOperationException("No se encontro DefaultConnection.");
    private readonly string? _masterKey = configuration["FacturacionElectronica:MasterKey"];

    public async Task<ComprobanteEmisionResultado> EmitirAsync(int negocioId, int comprobanteId, string usuario)
        => await EmitirCoreAsync(negocioId, comprobanteId, usuario, forzarManual: false);

    public async Task<ComprobanteEmisionResultado> EmitirManualAsync(int negocioId, int comprobanteId, string usuario)
        => await EmitirCoreAsync(negocioId, comprobanteId, usuario, forzarManual: true);

    public async Task<ComprobanteEmisionResultado> ConsultarEstadoAsync(int negocioId, int comprobanteId, string usuario)
    {
        var comprobante = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, comprobanteId);
        if (comprobante is null)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "NO_ENCONTRADO",
                Mensaje = "No se encontro el comprobante para emision."
            };
        }

        var config = await ObtenerConfiguracionAsync(negocioId);
        if (config is null)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "SIN_CONFIG",
                Mensaje = "No hay configuracion activa de proveedor."
            };
        }

        var esNubeFactConsulta = string.Equals(config.ProveedorCodigo, "NUBEFACT", StringComparison.OrdinalIgnoreCase);
        var esFeasyConsulta = string.Equals(config.ProveedorCodigo, "FEASY", StringComparison.OrdinalIgnoreCase);
        if (!esNubeFactConsulta && !esFeasyConsulta)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "PROVEEDOR_NO_IMPLEMENTADO",
                Mensaje = $"Proveedor SUNAT no implementado: {config.ProveedorCodigo}."
            };
        }

        object payload;
        if (esNubeFactConsulta)
        {
            payload = await ConstruirPayloadConsultaNubeFactAsync(negocioId, comprobanteId);
        }
        else
        {
            payload = await ConstruirPayloadConsultaFeasyAsync(negocioId, comprobanteId);
        }

        try
        {
            var endpointConsulta = esNubeFactConsulta
                ? config.BaseUrl
                : ResolverEndpointFeasyConsulta(config);
            using var request = new HttpRequestMessage(HttpMethod.Post, endpointConsulta);
            request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
            AplicarAutenticacion(request, config);

            var client = httpClientFactory.CreateClient();
            client.Timeout = TimeSpan.FromSeconds(config.TimeoutSegundos <= 0 ? 30 : config.TimeoutSegundos);
            using var response = await client.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var mensajeError = string.IsNullOrWhiteSpace(body) ? $"HTTP {(int)response.StatusCode}" : body;
                return new ComprobanteEmisionResultado
                {
                    Exito = false,
                    Codigo = $"HTTP_{(int)response.StatusCode}",
                    Mensaje = Truncar(mensajeError, 200) ?? "Error de comunicacion con proveedor."
                };
            }

            var respuesta = esNubeFactConsulta ? ParsearRespuestaNubeFact(body) : ParsearRespuestaFeasy(body);
            var comprobanteActual = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, comprobanteId);
            var debeForzarAceptado = await DebeForzarAceptadoPorPdfNubefactAsync(
                negocioId,
                comprobanteId,
                comprobanteActual?.CodigoDocumentoComprobante,
                esNubeFactConsulta,
                respuesta.UrlPdf);
            var esAceptadoFinal = respuesta.EsAceptado || debeForzarAceptado;
            var urlPrincipal = respuesta.UrlPdf ?? respuesta.UrlXml ?? respuesta.UrlCdr ?? respuesta.Url;
            var mensajeResultado = string.IsNullOrWhiteSpace(urlPrincipal) ? Truncar(respuesta.Mensaje, 500) : urlPrincipal;
            await RegistrarResultadoAsync(
                negocioId,
                comprobanteId,
                esAceptadoFinal ? 3 : 2,
                Truncar(respuesta.Codigo, 50),
                Truncar(mensajeResultado, 500),
                Truncar(respuesta.Ticket, 40),
                Truncar(respuesta.Hash, 100),
                Truncar(respuesta.UrlPdf, 500),
                Truncar(respuesta.UrlXml, 500),
                Truncar(respuesta.UrlCdr, 500),
                usuario);
            return new ComprobanteEmisionResultado
            {
                Exito = true,
                Codigo = respuesta.Codigo,
                Mensaje = respuesta.Mensaje
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error enviando comprobante {ComprobanteId} del negocio {NegocioId}.", comprobanteId, negocioId);
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "ERROR_ENVIO",
                Mensaje = "No se pudo enviar al proveedor electronico."
            };
        }
    }

    private async Task<ComprobanteEmisionResultado> EmitirCoreAsync(int negocioId, int comprobanteId, string usuario, bool forzarManual)
    {
        var comprobante = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, comprobanteId);
        if (comprobante is null)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "NO_ENCONTRADO",
                Mensaje = "No se encontro el comprobante para emision."
            };
        }

        if (!comprobante.EsTributario)
        {
            await RegistrarResultadoAsync(negocioId, comprobanteId, 3, "NO_TRIBUTARIO", "Recibo interno: no requiere proveedor electronico.", null, null, null, null, null, usuario);
            return new ComprobanteEmisionResultado
            {
                Exito = true,
                Codigo = "NO_TRIBUTARIO",
                Mensaje = "Comprobante no tributario, sin envio a proveedor."
            };
        }

        var config = await ObtenerConfiguracionAsync(negocioId);
        if (config is null)
        {
            await RegistrarResultadoAsync(negocioId, comprobanteId, 4, "SIN_CONFIG", "No hay configuracion activa de proveedor para este negocio.", null, null, null, null, null, usuario);
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "SIN_CONFIG",
                Mensaje = "No hay configuracion activa de proveedor."
            };
        }

        if (!config.EnviarComprobanteAutomatico && !forzarManual)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = true,
                Codigo = "AUTO_ENVIO_OFF",
                Mensaje = "El negocio tiene desactivado el envio automatico de comprobantes."
            };
        }
        var esNubeFact = string.Equals(config.ProveedorCodigo, "NUBEFACT", StringComparison.OrdinalIgnoreCase);
        var esFeasy = string.Equals(config.ProveedorCodigo, "FEASY", StringComparison.OrdinalIgnoreCase);
        if (!esNubeFact && !esFeasy)
        {
            await RegistrarResultadoAsync(
                negocioId,
                comprobanteId,
                1,
                "PROVEEDOR_NO_IMPLEMENTADO",
                $"Proveedor SUNAT no implementado: {config.ProveedorCodigo}. Comprobante pendiente de envio.",
                null,
                null,
                null,
                null,
                null,
                usuario);
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "PROVEEDOR_NO_IMPLEMENTADO",
                Mensaje = $"Proveedor SUNAT no implementado: {config.ProveedorCodigo}. El comprobante quedo pendiente."
            };
        }

        object payload;
        if (esNubeFact)
        {
            payload = await ConstruirPayloadNubeFactAsync(negocioId, comprobanteId, comprobante);
        }
        else
        {
            payload = await ConstruirPayloadFeasyAsync(negocioId, comprobanteId, comprobante, config);
        }

        try
        {
            var endpoint = esNubeFact
                ? config.BaseUrl
                : ResolverEndpointFeasy(config, comprobante.CodigoDocumentoComprobante);
            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
            request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
            AplicarAutenticacion(request, config);

            var client = httpClientFactory.CreateClient();
            client.Timeout = TimeSpan.FromSeconds(config.TimeoutSegundos <= 0 ? 30 : config.TimeoutSegundos);
            using var response = await client.SendAsync(request);
            var body = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var mensajeError = string.IsNullOrWhiteSpace(body) ? $"HTTP {(int)response.StatusCode}" : body;
                await RegistrarResultadoAsync(negocioId, comprobanteId, 4, $"HTTP_{(int)response.StatusCode}", Truncar(mensajeError, 500), null, null, null, null, null, usuario);
                return new ComprobanteEmisionResultado
                {
                    Exito = false,
                    Codigo = $"HTTP_{(int)response.StatusCode}",
                    Mensaje = Truncar(mensajeError, 200) ?? "Error de comunicacion con proveedor."
                };
            }

            var respuesta = esNubeFact ? ParsearRespuestaNubeFact(body) : ParsearRespuestaFeasy(body);
            var urlPrincipal = respuesta.UrlPdf ?? respuesta.UrlXml ?? respuesta.UrlCdr ?? respuesta.Url;
            var debeForzarAceptado = await DebeForzarAceptadoPorPdfNubefactAsync(
                negocioId,
                comprobanteId,
                comprobante.CodigoDocumentoComprobante,
                esNubeFact,
                respuesta.UrlPdf);
            var esAceptadoFinal = respuesta.EsAceptado || debeForzarAceptado;

            if (esFeasy && respuesta.EsAceptado && string.IsNullOrWhiteSpace(urlPrincipal))
            {
                // FEASY puede aceptar la emision y publicar URLs unos segundos despues; consultamos inmediatamente.
                await ConsultarEstadoAsync(negocioId, comprobanteId, usuario);
                var recargado = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, comprobanteId);
                urlPrincipal = recargado?.UrlDescargaProveedor;
            }

            var mensajeResultado = string.IsNullOrWhiteSpace(urlPrincipal) ? respuesta.Mensaje : urlPrincipal;
            await RegistrarResultadoAsync(
                negocioId,
                comprobanteId,
                esAceptadoFinal ? 3 : 2,
                Truncar(respuesta.Codigo, 50),
                Truncar(mensajeResultado, 500),
                Truncar(respuesta.Ticket, 40),
                Truncar(respuesta.Hash, 100),
                Truncar(respuesta.UrlPdf, 500),
                Truncar(respuesta.UrlXml, 500),
                Truncar(respuesta.UrlCdr, 500),
                usuario);

            var exitoOperacion = !string.IsNullOrWhiteSpace(respuesta.Codigo)
                                 && !respuesta.Codigo.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase)
                                 && !respuesta.Codigo.StartsWith("HTTP_", StringComparison.OrdinalIgnoreCase)
                                 && !respuesta.Mensaje.Contains("token incorrecto", StringComparison.OrdinalIgnoreCase)
                                 && !respuesta.Mensaje.Contains("acceso denegado", StringComparison.OrdinalIgnoreCase);

            return new ComprobanteEmisionResultado
            {
                Exito = exitoOperacion,
                Codigo = respuesta.Codigo,
                Mensaje = respuesta.Mensaje
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error enviando comprobante {ComprobanteId} del negocio {NegocioId}.", comprobanteId, negocioId);
            await RegistrarResultadoAsync(negocioId, comprobanteId, 4, "ERROR_ENVIO", Truncar(ex.Message, 500), null, null, null, null, null, usuario);
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "ERROR_ENVIO",
                Mensaje = "No se pudo enviar al proveedor electronico."
            };
        }
    }

    private async Task<object> ConstruirPayloadConsultaNubeFactAsync(int negocioId, int comprobanteId)
    {
        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, comprobanteId)
                             ?? throw new InvalidOperationException("No se pudo cargar el comprobante base para consulta.");
        return new Dictionary<string, object?>
        {
            ["operacion"] = "consultar_comprobante",
            ["tipo_de_comprobante"] = comprobanteBase.CodigoDocumentoComprobantenb,
            ["serie"] = comprobanteBase.Serie,
            ["numero"] = comprobanteBase.Numero
        };
    }

    private async Task<object> ConstruirPayloadConsultaFeasyAsync(int negocioId, int comprobanteId)
    {
        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, comprobanteId)
                             ?? throw new InvalidOperationException("No se pudo cargar el comprobante base para consulta.");

        var visual = await spService.ComprobantesObtenerVisualizacionAsync(negocioId, comprobanteId)
                     ?? throw new InvalidOperationException("No se pudo cargar la visualizacion del comprobante para consulta.");

        var codigoSunat = (visual.CodigoDocumentoComprobante ?? string.Empty).Trim();
        if (codigoSunat is not ("01" or "03" or "07" or "08"))
            throw new InvalidOperationException($"Tipo de documento no soportado para consulta FEASY: {codigoSunat}.");
        var (tipoDocumentoEmisor, numeroDocumentoEmisor) = SepararTipoNumeroDocumento((visual.NegocioDocumento ?? string.Empty).Trim(), "6");

        return new Dictionary<string, object?>
        {
            ["codigo_tipo_documento_emisor"] = tipoDocumentoEmisor,
            ["numero_documento_emisor"] = numeroDocumentoEmisor,
            ["codigo_tipo_documento"] = codigoSunat,
            ["serie_documento"] = (comprobanteBase.Serie ?? string.Empty).Trim().ToUpperInvariant(),
            ["numero_documento"] = comprobanteBase.Numero.ToString("D8", CultureInfo.InvariantCulture)
        };
    }

    private static string? LeerString(JsonDocument json, string propiedad)
    {
        return json.RootElement.TryGetProperty(propiedad, out var node) && node.ValueKind == JsonValueKind.String
            ? node.GetString()
            : null;
    }

    private async Task<bool> DebeForzarAceptadoPorPdfNubefactAsync(
        int negocioId,
        int comprobanteId,
        string? codigoDocumento,
        bool esNubefact,
        string? urlPdf)
    {
        if (!esNubefact || string.IsNullOrWhiteSpace(urlPdf))
            return false;

        var codigo = (codigoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        if (codigo == "03")
            return true;

        if (codigo != "07")
            return false;

        var actual = await spService.ComprobantesObtenerAsync(negocioId, comprobanteId);
        if (actual?.ComprobanteReferenciaId is null or <= 0)
            return false;

        var referencia = await spService.ComprobantesObtenerAsync(negocioId, actual.ComprobanteReferenciaId.Value);
        var codigoRef = (referencia?.CodigoDocumentoComprobante ?? string.Empty).Trim().ToUpperInvariant();
        return codigoRef == "03";
    }

    private void AplicarAutenticacion(HttpRequestMessage request, FacturacionConfig config)
    {
        string? Get(string key) => config.Credenciales.TryGetValue(key, out var value) ? value : null;
        static string NormalizarBearer(string raw)
        {
            var value = (raw ?? string.Empty).Trim();
            value = value.Trim('"', '\'');
            value = value.Replace("\0", string.Empty, StringComparison.Ordinal);
            if (value.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                value = value[7..].Trim();
            return value;
        }

        switch ((config.TipoAutenticacion ?? string.Empty).Trim().ToUpperInvariant())
        {
            case "API_KEY":
                var apiKey = Get("API_KEY");
                var tokenApiKey = Get("TOKEN");
                if (!string.IsNullOrWhiteSpace(tokenApiKey))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizarBearer(tokenApiKey));
                }
                else if (!string.IsNullOrWhiteSpace(apiKey))
                {
                    var apiKeyNormalizada = apiKey.Trim();
                    if (apiKeyNormalizada.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizarBearer(apiKeyNormalizada));
                    }
                    else if (string.Equals(config.ProveedorCodigo, "FEASY", StringComparison.OrdinalIgnoreCase))
                    {
                        // FEASY requiere token Bearer aun cuando la credencial este guardada en API_KEY.
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizarBearer(apiKeyNormalizada));
                    }
                    else
                    {
                        request.Headers.TryAddWithoutValidation("Authorization", apiKeyNormalizada);
                    }
                }
                break;
            case "TOKEN_FIJO":
                var token = Get("TOKEN");
                if (!string.IsNullOrWhiteSpace(token))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", NormalizarBearer(token));
                }
                break;
            case "USER_PASS":
                var user = Get("USUARIO");
                var pass = Get("PASSWORD");
                if (!string.IsNullOrWhiteSpace(user) && pass is not null)
                {
                    var raw = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{pass}"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Basic", raw);
                }
                break;
            case "OAUTH2_CLIENT":
                var clientId = Get("CLIENT_ID");
                var clientSecret = Get("CLIENT_SECRET");
                if (!string.IsNullOrWhiteSpace(clientId) && !string.IsNullOrWhiteSpace(clientSecret))
                {
                    var raw = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{clientId}:{clientSecret}"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Basic", raw);
                }
                break;
        }
    }

    private async Task<object> ConstruirPayloadFeasyAsync(int negocioId, int comprobanteId, ComprobanteVisualizacionViewModel comprobante, FacturacionConfig config)
    {
        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, comprobanteId)
                             ?? throw new InvalidOperationException("No se pudo cargar el comprobante base para emision.");

        var codigoSunat = (comprobante.CodigoDocumentoComprobante ?? string.Empty).Trim();
        if (codigoSunat is not ("01" or "03" or "07" or "08"))
            throw new InvalidOperationException($"Tipo de documento no soportado para FEASY: {codigoSunat}.");

        var clienteTipoDoc = (comprobanteBase.ClienteTipoDocumento ?? string.Empty).Trim();
        var clienteNumeroDoc = (comprobanteBase.ClienteNumeroDocumento ?? string.Empty).Trim();
        var clienteUbigeo = (comprobanteBase.ClienteCodigoUbigeo ?? string.Empty).Trim();
        var descripcion = $"ALQUILER DE ESPACIO DEPORTIVO - {comprobante.EspacioNombre} ({comprobante.FechaReserva:dd/MM/yyyy} {comprobante.HoraInicioReserva:HH\\:mm}-{comprobante.HoraFinReserva:HH\\:mm})";
        if (descripcion.Length > 250)
            descripcion = descripcion[..250];

        string? LeerCredencial(string key) => config.Credenciales.TryGetValue(key, out var value) ? value?.Trim() : null;

        var serie = (comprobante.Serie ?? string.Empty).Trim().ToUpperInvariant();
        var numeroDocumento = comprobante.Numero.ToString("D8", CultureInfo.InvariantCulture);
        var codigoInterno = $"{codigoSunat}{serie}{numeroDocumento}";
        var porcentajeIgv = decimal.Round(comprobante.PorcentajeIgv <= 0 ? 18m : comprobante.PorcentajeIgv, 2, MidpointRounding.AwayFromZero);
        var montoTotal = decimal.Round(comprobante.Total, 2, MidpointRounding.AwayFromZero);
        var montoIgv = decimal.Round(comprobante.Igv, 2, MidpointRounding.AwayFromZero);
        var montoGravado = decimal.Round(comprobante.SubTotal, 2, MidpointRounding.AwayFromZero);
        var cantidad = 1m;
        var montoValorUnitario = montoGravado;
        var montoPrecioUnitario = montoTotal;
        var codigoMoneda = comprobanteBase.MonedaNubefact == 2 ? "USD" : "PEN";
        var documentoEmisorRaw = (comprobante.NegocioDocumento ?? string.Empty).Trim();
        var (tipoDocumentoEmisor, numeroDocumentoEmisor) = SepararTipoNumeroDocumento(documentoEmisorRaw, "6");

        var informacionDocumento = new Dictionary<string, object?>
        {
            ["codigo_interno"] = codigoInterno,
            ["fecha_emision"] = comprobante.FechaEmision.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            ["hora_emision"] = comprobante.FechaEmision.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
            ["codigo_tipo_documento"] = codigoSunat,
            ["serie_documento"] = serie,
            ["numero_documento"] = numeroDocumento,
            ["observacion"] = $"Reserva {comprobante.ReservaId} - {comprobante.SedeNombre}".Trim(),
            ["correo"] = (comprobante.ClienteCorreo ?? string.Empty).Trim(),
            ["codigo_moneda"] = codigoMoneda,
            ["porcentaje_igv"] = porcentajeIgv,
            ["monto_total_gravado"] = montoGravado,
            ["monto_total_igv"] = montoIgv,
            ["monto_total"] = montoTotal
        };

        if (codigoSunat == "01")
        {
            informacionDocumento["fecha_vencimiento"] = comprobante.FechaEmision.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            informacionDocumento["forma_pago"] = "1";
        }

        var informacionEmisor = new Dictionary<string, object?>
        {
            ["codigo_tipo_documento_emisor"] = tipoDocumentoEmisor,
            ["numero_documento_emisor"] = numeroDocumentoEmisor,
            ["nombre_razon_social_emisor"] = (comprobante.NegocioRazonSocial ?? comprobante.NegocioNombre ?? string.Empty).Trim(),
            ["ubigeo_emisor"] = LeerCredencial("EMISOR_UBIGEO"),
            ["departamento_emisor"] = LeerCredencial("EMISOR_DEPARTAMENTO") ?? comprobante.NegocioDepartamento,
            ["provincia_emisor"] = LeerCredencial("EMISOR_PROVINCIA") ?? comprobante.NegocioProvincia,
            ["distrito_emisor"] = LeerCredencial("EMISOR_DISTRITO") ?? comprobante.NegocioDistrito,
            ["urbanizacion_emisor"] = LeerCredencial("EMISOR_URBANIZACION"),
            ["direccion_emisor"] = (comprobante.NegocioDireccionFiscal ?? string.Empty).Trim()
        };

        var informacionAdquiriente = new Dictionary<string, object?>
        {
            ["codigo_tipo_documento_adquiriente"] = string.IsNullOrWhiteSpace(clienteTipoDoc) ? "0" : clienteTipoDoc,
            ["numero_documento_adquiriente"] = clienteNumeroDoc,
            ["nombre_razon_social_adquiriente"] = (comprobante.ClienteNombre ?? string.Empty).Trim(),
            ["codigo_pais_adquiriente"] = "PE",
            ["ubigeo_adquiriente"] = clienteUbigeo,
            ["departamento_adquiriente"] = comprobante.ClienteDepartamento,
            ["provincia_adquiriente"] = comprobante.ClienteProvincia,
            ["distrito_adquiriente"] = comprobante.ClienteDistrito,
            ["urbanizacion_adquiriente"] = null,
            ["direccion_adquiriente"] = (comprobante.ClienteDireccion ?? string.Empty).Trim(),
            ["correo_adquiriente"] = (comprobante.ClienteCorreo ?? string.Empty).Trim()
        };

        var listaItems = new[]
        {
            new Dictionary<string, object?>
            {
                ["correlativo"] = 1,
                ["codigo_interno"] = "S001",
                ["codigo_sunat"] = null,
                ["tipo"] = "S",
                ["codigo_unidad_medida"] = "ZZ",
                ["descripcion"] = descripcion,
                ["cantidad"] = cantidad,
                ["monto_valor_unitario"] = decimal.Round(montoValorUnitario, 10, MidpointRounding.AwayFromZero),
                ["monto_precio_unitario"] = decimal.Round(montoPrecioUnitario, 10, MidpointRounding.AwayFromZero),
                ["monto_descuento"] = null,
                ["monto_valor_total"] = decimal.Round(montoGravado, 10, MidpointRounding.AwayFromZero),
                ["codigo_isc"] = null,
                ["monto_isc"] = null,
                ["codigo_indicador_afecto"] = "10",
                ["monto_igv"] = montoIgv,
                ["monto_impuesto_bolsa"] = null,
                ["monto_total"] = montoTotal
            }
        };

        var payload = new Dictionary<string, object?>
        {
            ["informacion_documento"] = informacionDocumento,
            ["informacion_emisor"] = informacionEmisor,
            ["informacion_adquiriente"] = informacionAdquiriente,
            ["lista_items"] = listaItems
        };

        if (codigoSunat == "01")
        {
            payload["informacion_entrega_bienes"] = new Dictionary<string, object?>
            {
                ["codigo_pais_entrega"] = "PE",
                ["ubigeo_entrega"] = clienteUbigeo,
                ["departamento_entrega"] = comprobante.ClienteDepartamento,
                ["provincia_entrega"] = comprobante.ClienteProvincia,
                ["distrito_entrega"] = comprobante.ClienteDistrito,
                ["urbanizacion_entrega"] = null,
                ["direccion_entrega"] = (comprobante.ClienteDireccion ?? string.Empty).Trim()
            };
            payload["indicadores"] = new Dictionary<string, object?>
            {
                ["indicador_entrega_bienes"] = false
            };
        }

        if (codigoSunat is "07" or "08")
        {
            if (!comprobanteBase.ComprobanteReferenciaId.HasValue)
                throw new InvalidOperationException("NC/ND requiere comprobante de referencia.");

            var referencia = await spService.ComprobantesObtenerAsync(negocioId, comprobanteBase.ComprobanteReferenciaId.Value)
                             ?? throw new InvalidOperationException("No se encontro comprobante de referencia para NC/ND.");
            var referenciaCodigoSunat = (referencia.CodigoDocumentoComprobante ?? string.Empty).Trim();
            if (referenciaCodigoSunat is not ("01" or "03"))
                throw new InvalidOperationException($"Documento referencia no valido para NC/ND FEASY: {referenciaCodigoSunat}.");

            informacionDocumento["forma_pago"] = "1";
            var codigoNota = (comprobanteBase.TipoNotaCodigoSunat ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(codigoNota))
                throw new InvalidOperationException("NC/ND requiere TipoNotaCodigoSunat.");

            if (codigoSunat == "07")
                informacionDocumento["codigo_tipo_nota_credito"] = codigoNota;
            else
                informacionDocumento["codigo_tipo_nota_debito"] = codigoNota;

            payload["informacion_documento_referencia"] = new Dictionary<string, object?>
            {
                ["codigo_tipo_documento_referencia"] = referenciaCodigoSunat,
                ["serie_documento_referencia"] = (referencia.Serie ?? string.Empty).Trim().ToUpperInvariant(),
                ["numero_documento_referencia"] = referencia.Numero.ToString("D8", CultureInfo.InvariantCulture)
            };
        }

        return payload;
    }

    private async Task<FacturacionConfig?> ObtenerConfiguracionAsync(int negocioId)
    {
        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_FacturacionProveedorConfig_ObtenerPorNegocio", cn)
        {
            CommandType = CommandType.StoredProcedure
        };
        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;
        cmd.Parameters.Add("@Ambiente", SqlDbType.NVarChar, 15).Value = "PRODUCCION";

        await using var dr = await cmd.ExecuteReaderAsync();
        var configs = new List<FacturacionConfig>();
        while (await dr.ReadAsync())
        {
            configs.Add(new FacturacionConfig
            {
                ConfigId = dr.GetInt32(0),
                BaseUrl = dr.IsDBNull(7) ? string.Empty : dr.GetString(7),
                ProveedorCodigo = dr.IsDBNull(3) ? string.Empty : dr.GetString(3),
                TipoAutenticacion = dr.IsDBNull(5) ? string.Empty : dr.GetString(5),
                TimeoutSegundos = dr.IsDBNull(9) ? 30 : dr.GetInt32(9),
                EsDefault = !dr.IsDBNull(10) && dr.GetBoolean(10)
            });
        }

        if (!configs.Any())
            return null;

        await dr.NextResultAsync();
        while (await dr.ReadAsync())
        {
            var configId = dr.GetInt32(0);
            var tipo = dr.IsDBNull(1) ? string.Empty : dr.GetString(1);
            var bytes = (byte[])dr["SecretoCifrado"];
            var keyVersion = dr.IsDBNull(3) ? null : dr.GetString(3);
            var valor = Descifrar(bytes, keyVersion);

            var cfg = configs.FirstOrDefault(x => x.ConfigId == configId);
            if (cfg is not null && !string.IsNullOrWhiteSpace(tipo))
            {
                cfg.Credenciales[tipo.Trim().ToUpperInvariant()] = valor;
            }
        }

        var configFinal = configs.OrderByDescending(x => x.EsDefault).ThenBy(x => x.ConfigId).First();

        await using var cmdNegocio = new SqlCommand(
            "SELECT TOP (1) EnviarComprobanteAutomatico FROM dbo.Negocios WHERE Id = @NegocioId;",
            cn);
        cmdNegocio.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;
        var autoEnvio = await cmdNegocio.ExecuteScalarAsync();
        configFinal.EnviarComprobanteAutomatico = autoEnvio is not null && autoEnvio != DBNull.Value && Convert.ToBoolean(autoEnvio, CultureInfo.InvariantCulture);

        return configFinal;
    }

    private async Task<object> ConstruirPayloadNubeFactAsync(int negocioId, int comprobanteId, ComprobanteVisualizacionViewModel comprobante)
    {
        var comprobanteBase = await spService.ComprobantesObtenerAsync(negocioId, comprobanteId)
                             ?? throw new InvalidOperationException("No se pudo cargar el comprobante base para emision.");

        var tipoDocumento = comprobanteBase.CodigoDocumentoComprobantenb;
        if (tipoDocumento == 0)
            throw new InvalidOperationException($"Tipo de documento no soportado para NubeFact (CodigoDocumentoComprobantenb): {comprobanteBase.CodigoDocumentoComprobantenb}");

        var clienteTipoDoc = (comprobanteBase.ClienteTipoDocumento ?? string.Empty).Trim();
        var clienteNumeroDoc = (comprobanteBase.ClienteNumeroDocumento ?? string.Empty).Trim();
        var clienteUbigeo = (comprobanteBase.ClienteCodigoUbigeo ?? string.Empty).Trim();
        var descripcion = $"ALQUILER DE ESPACIO DEPORTIVO - {comprobante.EspacioNombre} ({comprobante.FechaReserva:dd/MM/yyyy} {comprobante.HoraInicioReserva:HH\\:mm}-{comprobante.HoraFinReserva:HH\\:mm})";
        var precioUnitario = comprobante.Total;
        var valorUnitario = comprobante.SubTotal;
        var igv = comprobante.Igv;
        object tipoDocClientePayload = string.IsNullOrWhiteSpace(clienteTipoDoc) ? "-" : clienteTipoDoc;

        var payload = new Dictionary<string, object?>
        {
            ["operacion"] = "generar_comprobante",
            ["tipo_de_comprobante"] = tipoDocumento,
            ["serie"] = comprobante.Serie,
            ["numero"] = comprobante.Numero,
            ["sunat_transaction"] = 1,
            ["cliente_tipo_de_documento"] = tipoDocClientePayload,
            ["cliente_numero_de_documento"] = clienteNumeroDoc,
            ["cliente_denominacion"] = (comprobante.ClienteNombre ?? string.Empty).Trim(),
            ["cliente_direccion"] = (comprobante.ClienteDireccion ?? string.Empty).Trim(),
            ["cliente_email"] = (comprobante.ClienteCorreo ?? string.Empty).Trim(),
            ["cliente_ubigeo"] = clienteUbigeo,
            ["cliente_email_1"] = "",
            ["cliente_email_2"] = "",
            ["fecha_de_emision"] = comprobante.FechaEmision.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture),
            ["fecha_de_vencimiento"] = "",
            ["moneda"] = comprobanteBase.MonedaNubefact <= 0 ? 1 : comprobanteBase.MonedaNubefact,
            ["tipo_de_cambio"] = "",
            ["porcentaje_de_igv"] = decimal.Round(comprobante.PorcentajeIgv, 2, MidpointRounding.AwayFromZero),
            ["descuento_global"] = "",
            ["total_descuento"] = "",
            ["total_anticipo"] = "",
            ["total_gravada"] = decimal.Round(comprobante.SubTotal, 2, MidpointRounding.AwayFromZero),
            ["total_inafecta"] = "",
            ["total_exonerada"] = "",
            ["total_igv"] = decimal.Round(comprobante.Igv, 2, MidpointRounding.AwayFromZero),
            ["total_gratuita"] = "",
            ["total_otros_cargos"] = "",
            ["total"] = decimal.Round(comprobante.Total, 2, MidpointRounding.AwayFromZero),
            ["percepcion_tipo"] = "",
            ["percepcion_base_imponible"] = "",
            ["total_percepcion"] = "",
            ["total_incluido_percepcion"] = "",
            ["detraccion"] = false,
            ["observaciones"] = $"Reserva {comprobante.ReservaId} - {comprobante.SedeNombre}".Trim(),
            ["documento_que_se_modifica_tipo"] = "",
            ["documento_que_se_modifica_serie"] = "",
            ["documento_que_se_modifica_numero"] = "",
            ["tipo_de_nota_de_credito"] = "",
            ["tipo_de_nota_de_debito"] = "",
            ["enviar_automaticamente_a_la_sunat"] = true,
            ["enviar_automaticamente_al_cliente"] = false,
            ["codigo_unico"] = "",
            ["condiciones_de_pago"] = "",
            ["medio_de_pago"] = "",
            ["placa_vehiculo"] = "",
            ["orden_compra_servicio"] = "",
            ["tabla_personalizada_codigo"] = "",
            ["formato_de_pdf"] = "",
            ["items"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["unidad_de_medida"] = "ZZ",
                    ["codigo"] = "001",
                    ["descripcion"] = descripcion.Length > 250 ? descripcion[..250] : descripcion,
                    ["cantidad"] = 1m,
                    ["valor_unitario"] = decimal.Round(valorUnitario, 2, MidpointRounding.AwayFromZero),
                    ["precio_unitario"] = decimal.Round(precioUnitario, 2, MidpointRounding.AwayFromZero),
                    ["descuento"] = 0m,
                    ["subtotal"] = decimal.Round(comprobante.SubTotal, 2, MidpointRounding.AwayFromZero),
                    ["tipo_de_igv"] = 1,
                    ["igv"] = decimal.Round(igv, 2, MidpointRounding.AwayFromZero),
                    ["total"] = decimal.Round(comprobante.Total, 2, MidpointRounding.AwayFromZero),
                    ["anticipo_regularizacion"] = false,
                    ["anticipo_documento_serie"] = string.Empty,
                    ["anticipo_documento_numero"] = string.Empty
                    ,
                    ["codigo_producto_sunat"] = "20000000"
                }
            }
        };

        if (tipoDocumento is 3 or 4)
        {
            if (!comprobanteBase.ComprobanteReferenciaId.HasValue)
                throw new InvalidOperationException("NC/ND requiere comprobante de referencia.");

            var referencia = await spService.ComprobantesObtenerAsync(negocioId, comprobanteBase.ComprobanteReferenciaId.Value)
                             ?? throw new InvalidOperationException("No se encontro comprobante de referencia para NC/ND.");

            if (referencia.CodigoDocumentoComprobantenb == 0)
                throw new InvalidOperationException("El comprobante de referencia no tiene CodigoDocumentoComprobantenb valido.");

            payload["documento_que_se_modifica_tipo"] = referencia.CodigoDocumentoComprobantenb;
            payload["documento_que_se_modifica_serie"] = referencia.Serie;
            payload["documento_que_se_modifica_numero"] = referencia.Numero;

            var codigoNota = (comprobanteBase.TipoNotaCodigoSunat ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(codigoNota))
                throw new InvalidOperationException("NC/ND requiere TipoNotaCodigoSunat.");

            if (tipoDocumento == 3)
                payload["tipo_de_nota_de_credito"] = Convert.ToInt32(codigoNota, CultureInfo.InvariantCulture);
            else
                payload["tipo_de_nota_de_debito"] = Convert.ToInt32(codigoNota, CultureInfo.InvariantCulture);
        }

        return payload;
    }

    private static RespuestaNubeFact ParsearRespuestaNubeFact(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return new RespuestaNubeFact
            {
                Codigo = "SIN_RESPUESTA",
                Mensaje = "Proveedor no devolvio contenido."
            };
        }

        using var json = JsonDocument.Parse(body);
        var root = json.RootElement;

        var errores = new List<string>();
        if (root.TryGetProperty("errors", out var errorsNode) && errorsNode.ValueKind == JsonValueKind.Array)
        {
            foreach (var err in errorsNode.EnumerateArray())
            {
                if (err.ValueKind == JsonValueKind.String)
                {
                    var txt = err.GetString();
                    if (!string.IsNullOrWhiteSpace(txt))
                        errores.Add(txt.Trim());
                }
                else if (err.ValueKind == JsonValueKind.Object)
                {
                    var m = err.TryGetProperty("message", out var mm) && mm.ValueKind == JsonValueKind.String ? mm.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(m))
                        errores.Add(m.Trim());
                }
            }
        }

        var mensaje = LeerString(json, "mensaje")
                      ?? LeerString(json, "message")
                      ?? (errores.Count > 0 ? string.Join(" | ", errores) : "Comprobante enviado.");

        var codigo = LeerString(json, "codigo")
                     ?? LeerString(json, "codigo_error")
                     ?? (errores.Count > 0 ? "ERROR_VALIDACION" : "ENVIADO");

        var aceptado = false;
        if (root.TryGetProperty("aceptada_por_sunat", out var aceptadaNode))
        {
            aceptado = aceptadaNode.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Number => aceptadaNode.GetInt32() == 1,
                JsonValueKind.String => string.Equals(aceptadaNode.GetString(), "true", StringComparison.OrdinalIgnoreCase)
                                        || aceptadaNode.GetString() == "1",
                _ => false
            };
        }

        return new RespuestaNubeFact
        {
            Codigo = codigo,
            Mensaje = mensaje,
            Ticket = LeerString(json, "sunat_ticket_numero") ?? LeerString(json, "ticket"),
            Hash = LeerString(json, "cadena_para_codigo_qr") ?? LeerString(json, "hash"),
            Url = LeerString(json, "enlace"),
            UrlPdf = LeerString(json, "enlace_del_pdf") ?? LeerString(json, "url_pdf"),
            UrlXml = LeerString(json, "enlace_del_xml") ?? LeerString(json, "url_xml"),
            UrlCdr = LeerString(json, "enlace_del_cdr") ?? LeerString(json, "url_cdr"),
            EsAceptado = aceptado
        };
    }

    private static RespuestaNubeFact ParsearRespuestaFeasy(string body)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return new RespuestaNubeFact
            {
                Codigo = "SIN_RESPUESTA",
                Mensaje = "Proveedor no devolvio contenido."
            };
        }

        using var json = JsonDocument.Parse(body);
        var root = json.RootElement;

        string? BuscarTexto(JsonElement node, params string[] keys)
        {
            foreach (var key in keys)
            {
                if (node.TryGetProperty(key, out var child) && child.ValueKind == JsonValueKind.String)
                {
                    var value = child.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }
            return null;
        }

        JsonElement? dataNode = null;
        if (root.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.Object)
            dataNode = data;
        else if (root.TryGetProperty("resultado", out var resultado) && resultado.ValueKind == JsonValueKind.Object)
            dataNode = resultado;

        var mensaje = (dataNode.HasValue ? BuscarTexto(dataNode.Value, "mensaje_respuesta", "mensaje", "message") : null)
                      ?? BuscarTexto(root, "mensaje", "message")
                      ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "mensaje", "message") : null)
                      ?? "Comprobante enviado.";
        var codigo = (dataNode.HasValue ? BuscarTexto(dataNode.Value, "codigo_respuesta", "codigo", "code") : null)
                     ?? BuscarTexto(root, "codigo", "code", "status_code")
                     ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "codigo", "code") : null);
        var ticket = BuscarTexto(root, "ticket", "numero_ticket", "sunat_ticket")
                     ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "ticket", "numero_ticket", "sunat_ticket") : null);
        var hash = BuscarTexto(root, "hash", "codigo_hash", "hash_cpe")
                   ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "hash", "codigo_hash", "hash_cpe") : null);
        var urlPdf = BuscarTexto(root, "url_pdf", "enlace_pdf", "pdf")
                     ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "url_pdf", "enlace_pdf", "pdf", "ruta_reporte") : null);
        var urlXml = BuscarTexto(root, "url_xml", "enlace_xml", "xml")
                     ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "url_xml", "enlace_xml", "xml", "ruta_xml") : null);
        var urlCdr = BuscarTexto(root, "url_cdr", "enlace_cdr", "cdr")
                     ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "url_cdr", "enlace_cdr", "cdr", "ruta_cdr") : null);
        var url = BuscarTexto(root, "url", "enlace")
                  ?? (dataNode.HasValue ? BuscarTexto(dataNode.Value, "url", "enlace") : null);

        var exito = true;
        if (root.TryGetProperty("success", out var successNode))
        {
            exito = successNode.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Number => successNode.GetInt32() == 1,
                JsonValueKind.String => string.Equals(successNode.GetString(), "true", StringComparison.OrdinalIgnoreCase)
                                        || successNode.GetString() == "1",
                _ => true
            };
        }
        else if (root.TryGetProperty("aceptada_por_sunat", out var aceptadaNode))
        {
            exito = aceptadaNode.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Number => aceptadaNode.GetInt32() == 1,
                JsonValueKind.String => string.Equals(aceptadaNode.GetString(), "true", StringComparison.OrdinalIgnoreCase)
                                        || aceptadaNode.GetString() == "1",
                _ => true
            };
        }
        else if (mensaje.Contains("error", StringComparison.OrdinalIgnoreCase))
        {
            exito = false;
        }

        codigo ??= exito ? "ENVIADO" : "ERROR_FEASY";

        return new RespuestaNubeFact
        {
            Codigo = codigo,
            Mensaje = mensaje,
            Ticket = ticket,
            Hash = hash,
            Url = url,
            UrlPdf = urlPdf,
            UrlXml = urlXml,
            UrlCdr = urlCdr,
            EsAceptado = exito
        };
    }

    private static string ResolverEndpointFeasy(FacturacionConfig config, string? codigoDocumentoSunat)
    {
        var baseUrl = (config.BaseUrl ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(baseUrl))
            throw new InvalidOperationException("FEASY requiere BaseUrl configurada.");

        var codigo = (codigoDocumentoSunat ?? string.Empty).Trim().ToUpperInvariant();
        var endpoint = codigo switch
        {
            "01" => "comprobante/enviar_factura",
            "03" => "comprobante/enviar_boleta",
            "07" => "comprobante/enviar_nota_credito",
            "08" => "comprobante/enviar_nota_debito",
            _ => throw new InvalidOperationException($"FEASY no soporta el tipo de documento {codigo}.")
        };

        return $"{baseUrl.TrimEnd('/')}/{endpoint}";
    }

    private static string ResolverEndpointFeasyConsulta(FacturacionConfig config)
    {
        var baseUrl = (config.BaseUrl ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(baseUrl))
            throw new InvalidOperationException("FEASY requiere BaseUrl configurada.");

        return $"{baseUrl.TrimEnd('/')}/comprobante/consultar";
    }

    private static (string Tipo, string Numero) SepararTipoNumeroDocumento(string? valor, string tipoPorDefecto)
    {
        var raw = (valor ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(raw))
            return (tipoPorDefecto, string.Empty);

        var idx = raw.IndexOf('-', StringComparison.Ordinal);
        if (idx > 0 && idx < raw.Length - 1)
        {
            var tipo = raw[..idx].Trim();
            var numero = raw[(idx + 1)..].Trim();
            if (!string.IsNullOrWhiteSpace(tipo) && !string.IsNullOrWhiteSpace(numero))
                return (tipo, numero);
        }

        return (tipoPorDefecto, raw);
    }

    private string Descifrar(byte[] secretoCifrado, string? keyVersion)
    {
        if (secretoCifrado.Length == 0)
            return string.Empty;

        if (string.IsNullOrWhiteSpace(_masterKey))
            return Encoding.UTF8.GetString(secretoCifrado);

        try
        {
            var keyMaterial = SHA256.HashData(Encoding.UTF8.GetBytes(_masterKey + "|" + (keyVersion ?? "v1")));
            var iv = secretoCifrado[..16];
            var cipher = secretoCifrado[16..];
            using var aes = Aes.Create();
            aes.Key = keyMaterial;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;
            using var decryptor = aes.CreateDecryptor();
            var plainBytes = decryptor.TransformFinalBlock(cipher, 0, cipher.Length);
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch
        {
            return Encoding.UTF8.GetString(secretoCifrado);
        }
    }

    private async Task RegistrarResultadoAsync(
        int negocioId,
        int comprobanteId,
        int estado,
        string? codigo,
        string? mensaje,
        string? ticket,
        string? hash,
        string? urlPdf,
        string? urlXml,
        string? urlCdr,
        string usuario)
    {
        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync();
        await using var cmd = new SqlCommand("Sp_Comprobantes_RegistrarEnvioProveedor", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        cmd.Parameters.Add("@NegocioId", SqlDbType.Int).Value = negocioId;
        cmd.Parameters.Add("@Id", SqlDbType.Int).Value = comprobanteId;
        cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = estado;
        cmd.Parameters.Add("@CodigoRespuesta", SqlDbType.NVarChar, 50).Value = (object?)codigo ?? DBNull.Value;
        cmd.Parameters.Add("@MensajeRespuesta", SqlDbType.NVarChar, 500).Value = (object?)mensaje ?? DBNull.Value;
        cmd.Parameters.Add("@NumeroTicketSunat", SqlDbType.NVarChar, 40).Value = (object?)ticket ?? DBNull.Value;
        cmd.Parameters.Add("@CodigoHashCpe", SqlDbType.NVarChar, 100).Value = (object?)hash ?? DBNull.Value;
        cmd.Parameters.Add("@UrlPdfSunat", SqlDbType.NVarChar, 500).Value = (object?)urlPdf ?? DBNull.Value;
        cmd.Parameters.Add("@UrlXmlSunat", SqlDbType.NVarChar, 500).Value = (object?)urlXml ?? DBNull.Value;
        cmd.Parameters.Add("@UrlCdrSunat", SqlDbType.NVarChar, 500).Value = (object?)urlCdr ?? DBNull.Value;
        cmd.Parameters.Add("@Usuario", SqlDbType.NVarChar, 200).Value = usuario;
        await cmd.ExecuteNonQueryAsync();
    }

    private static string? Truncar(string? texto, int max)
    {
        if (string.IsNullOrWhiteSpace(texto))
            return texto;

        return texto.Length <= max ? texto : texto[..max];
    }

    private sealed class FacturacionConfig
    {
        public int ConfigId { get; init; }
        public string ProveedorCodigo { get; init; } = string.Empty;
        public string BaseUrl { get; init; } = string.Empty;
        public string TipoAutenticacion { get; init; } = string.Empty;
        public int TimeoutSegundos { get; init; }
        public bool EsDefault { get; init; }
        public bool EnviarComprobanteAutomatico { get; set; }
        public Dictionary<string, string> Credenciales { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed class RespuestaNubeFact
    {
        public string Codigo { get; init; } = "ENVIADO";
        public string Mensaje { get; init; } = "Comprobante enviado.";
        public string? Ticket { get; init; }
        public string? Hash { get; init; }
        public string? Url { get; init; }
        public string? UrlPdf { get; init; }
        public string? UrlXml { get; init; }
        public string? UrlCdr { get; init; }
        public bool EsAceptado { get; init; }
    }
}
