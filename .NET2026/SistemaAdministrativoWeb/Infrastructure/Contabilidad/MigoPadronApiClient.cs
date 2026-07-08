using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoPadronApiClient(
    HttpClient httpClient,
    IOptions<MigoApiSettings> settings) : IMigoPadronApiClient
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public async Task<MigoRucResultDto?> ConsultarRucAsync(string ruc, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        ValidarConfiguracion(cfg);

        using var request = CrearRequest(
            cfg.RucPath,
            new Dictionary<string, string>
            {
                ["token"] = cfg.Token.Trim(),
                ["ruc"] = NormalizarSoloDigitos(ruc)
            });

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"La API de Migo devolvio HTTP {(int)response.StatusCode} al consultar el RUC.");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<MigoRucResponse>(stream, JsonOptions, cancellationToken);
        if (payload is null || !payload.Success || string.IsNullOrWhiteSpace(payload.Ruc))
        {
            return null;
        }

        return new MigoRucResultDto
        {
            Ruc = payload.Ruc.Trim(),
            NombreORazonSocial = (payload.NombreORazonSocial ?? string.Empty).Trim(),
            Ubigeo = NormalizarTexto(payload.Ubigeo),
            DireccionSimple = NormalizarTexto(payload.DireccionSimple),
            Direccion = NormalizarTexto(payload.Direccion),
            Distrito = NormalizarTexto(payload.Distrito),
            Provincia = NormalizarTexto(payload.Provincia),
            Departamento = NormalizarTexto(payload.Departamento),
            EstadoContribuyente = NormalizarTexto(payload.EstadoDelContribuyente),
            CondicionDomicilio = NormalizarTexto(payload.CondicionDeDomicilio)
        };
    }

    public async Task<MigoDniResultDto?> ConsultarDniAsync(string dni, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        ValidarConfiguracion(cfg);

        using var request = CrearRequest(
            cfg.DniPath,
            new Dictionary<string, string>
            {
                ["token"] = cfg.Token.Trim(),
                ["dni"] = NormalizarSoloDigitos(dni)
            });

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"La API de Migo devolvio HTTP {(int)response.StatusCode} al consultar el DNI.");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<MigoDniResponse>(stream, JsonOptions, cancellationToken);
        if (payload is null || !payload.Success || string.IsNullOrWhiteSpace(payload.Dni) || string.IsNullOrWhiteSpace(payload.Nombre))
        {
            return null;
        }

        return new MigoDniResultDto
        {
            Dni = payload.Dni.Trim(),
            NombreCompleto = payload.Nombre.Trim()
        };
    }

    public async Task<MigoCpeResultDto?> ValidarCpeAsync(MigoCpeRequestDto requestData, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        ValidarConfiguracion(cfg);

        var body = new Dictionary<string, string>
        {
            ["token"] = cfg.Token.Trim(),
            ["ruc_emisor"] = NormalizarSoloDigitos(requestData.RucEmisor),
            ["tipo_comprobante"] = (requestData.TipoComprobante ?? string.Empty).Trim().ToUpperInvariant(),
            ["serie"] = (requestData.Serie ?? string.Empty).Trim().ToUpperInvariant(),
            ["numero"] = (requestData.Numero ?? string.Empty).Trim().ToUpperInvariant(),
            ["fecha_emision"] = requestData.FechaEmisionTexto
        };

        if (!string.IsNullOrWhiteSpace(requestData.MontoTexto))
        {
            body["monto"] = requestData.MontoTexto!;
        }

        using var request = CrearRequest(cfg.CpePath, body);
        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"La API de Migo devolvio HTTP {(int)response.StatusCode} al validar el CPE.");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<MigoCpeResponse>(stream, JsonOptions, cancellationToken);
        if (payload is null)
        {
            return null;
        }

        return new MigoCpeResultDto
        {
            Success = payload.Success,
            EstadoComprobante = NormalizarTexto(payload.EstadoComprobante),
            EstadoContribuyente = NormalizarTexto(payload.EstadoContribuyente),
            CondicionDomicilio = NormalizarTexto(payload.CondicionDomicilio),
            Observaciones = NormalizarObservaciones(payload.Observaciones)
        };
    }

    private static void ValidarConfiguracion(MigoApiSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.BaseUrl) || string.IsNullOrWhiteSpace(settings.Token))
        {
            throw new InvalidOperationException("La configuracion de Migo API no esta completa. Revise la seccion MigoApi en los archivos de variables.");
        }
    }

    private HttpRequestMessage CrearRequest(string relativePath, IReadOnlyDictionary<string, string> body)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, relativePath)
        {
            Content = new StringContent(JsonSerializer.Serialize(body, JsonOptions), Encoding.UTF8, "application/json")
        };

        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        return request;
    }

    private static string NormalizarSoloDigitos(string value)
    {
        return new string((value ?? string.Empty).Where(char.IsDigit).ToArray());
    }

    private static string? NormalizarTexto(string? value)
    {
        var texto = string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        return string.Equals(texto, "-", StringComparison.Ordinal) ? null : texto;
    }

    private static string? NormalizarObservaciones(JsonElement observaciones)
    {
        if (observaciones.ValueKind == JsonValueKind.Undefined || observaciones.ValueKind == JsonValueKind.Null)
        {
            return null;
        }

        if (observaciones.ValueKind == JsonValueKind.String)
        {
            return NormalizarTexto(observaciones.GetString());
        }

        if (observaciones.ValueKind == JsonValueKind.Array)
        {
            var items = new List<string>();
            foreach (var item in observaciones.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                {
                    var texto = NormalizarTexto(item.GetString());
                    if (!string.IsNullOrWhiteSpace(texto))
                    {
                        items.Add(texto);
                    }
                }
                else if (item.ValueKind != JsonValueKind.Null && item.ValueKind != JsonValueKind.Undefined)
                {
                    items.Add(item.ToString());
                }
            }

            return items.Count == 0 ? null : string.Join(" | ", items);
        }

        return NormalizarTexto(observaciones.ToString());
    }

    private sealed class MigoRucResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("ruc")]
        public string? Ruc { get; set; }

        [JsonPropertyName("nombre_o_razon_social")]
        public string? NombreORazonSocial { get; set; }

        [JsonPropertyName("estado_del_contribuyente")]
        public string? EstadoDelContribuyente { get; set; }

        [JsonPropertyName("condicion_de_domicilio")]
        public string? CondicionDeDomicilio { get; set; }

        [JsonPropertyName("ubigeo")]
        public string? Ubigeo { get; set; }

        [JsonPropertyName("distrito")]
        public string? Distrito { get; set; }

        [JsonPropertyName("provincia")]
        public string? Provincia { get; set; }

        [JsonPropertyName("departamento")]
        public string? Departamento { get; set; }

        [JsonPropertyName("direccion_simple")]
        public string? DireccionSimple { get; set; }

        [JsonPropertyName("direccion")]
        public string? Direccion { get; set; }
    }

    private sealed class MigoDniResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("dni")]
        public string? Dni { get; set; }

        [JsonPropertyName("nombre")]
        public string? Nombre { get; set; }
    }

    private sealed class MigoCpeResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("estado_comprobante")]
        public string? EstadoComprobante { get; set; }

        [JsonPropertyName("estado_contribuyente")]
        public string? EstadoContribuyente { get; set; }

        [JsonPropertyName("condicion_domicilio")]
        public string? CondicionDomicilio { get; set; }

        [JsonPropertyName("observaciones")]
        public JsonElement Observaciones { get; set; }
    }
}
