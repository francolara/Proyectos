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

        if (!config.EnviarComprobanteAutomatico)
        {
            return new ComprobanteEmisionResultado
            {
                Exito = true,
                Codigo = "AUTO_ENVIO_OFF",
                Mensaje = "El negocio tiene desactivado el envio automatico de comprobantes."
            };
        }
        if (!string.Equals(config.ProveedorCodigo, "NUBEFACT", StringComparison.OrdinalIgnoreCase))
        {
            return new ComprobanteEmisionResultado
            {
                Exito = false,
                Codigo = "PROVEEDOR_NO_IMPLEMENTADO",
                Mensaje = $"Proveedor SUNAT no implementado: {config.ProveedorCodigo}."
            };
        }

        var payload = await ConstruirPayloadConsultaNubeFactAsync(negocioId, comprobanteId);

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, config.BaseUrl);
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

            var respuesta = ParsearRespuestaNubeFact(body);
            var urlPrincipal = respuesta.UrlPdf ?? respuesta.UrlXml ?? respuesta.UrlCdr ?? respuesta.Url;
            var mensajeResultado = string.IsNullOrWhiteSpace(urlPrincipal) ? Truncar(respuesta.Mensaje, 500) : urlPrincipal;
            await RegistrarResultadoAsync(
                negocioId,
                comprobanteId,
                respuesta.EsAceptado ? 3 : 2,
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
        if (!string.Equals(config.ProveedorCodigo, "NUBEFACT", StringComparison.OrdinalIgnoreCase))
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

        var payload = await ConstruirPayloadNubeFactAsync(negocioId, comprobanteId, comprobante);

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, config.BaseUrl);
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

            var respuesta = ParsearRespuestaNubeFact(body);
            var urlPrincipal = respuesta.UrlPdf ?? respuesta.UrlXml ?? respuesta.UrlCdr ?? respuesta.Url;
            var mensajeResultado = string.IsNullOrWhiteSpace(urlPrincipal) ? respuesta.Mensaje : urlPrincipal;
            await RegistrarResultadoAsync(
                negocioId,
                comprobanteId,
                respuesta.EsAceptado ? 3 : 2,
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

    private static string? LeerString(JsonDocument json, string propiedad)
    {
        return json.RootElement.TryGetProperty(propiedad, out var node) && node.ValueKind == JsonValueKind.String
            ? node.GetString()
            : null;
    }

    private void AplicarAutenticacion(HttpRequestMessage request, FacturacionConfig config)
    {
        string? Get(string key) => config.Credenciales.TryGetValue(key, out var value) ? value : null;

        switch ((config.TipoAutenticacion ?? string.Empty).Trim().ToUpperInvariant())
        {
            case "API_KEY":
                var apiKey = Get("API_KEY");
                if (!string.IsNullOrWhiteSpace(apiKey))
                {
                    request.Headers.TryAddWithoutValidation("Authorization", apiKey.Trim());
                }
                break;
            case "TOKEN_FIJO":
                var token = Get("TOKEN");
                if (!string.IsNullOrWhiteSpace(token))
                {
                    request.Headers.TryAddWithoutValidation("Authorization", token.Trim());
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
