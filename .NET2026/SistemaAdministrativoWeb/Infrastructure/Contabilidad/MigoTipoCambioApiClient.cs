using System.Globalization;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MigoTipoCambioApiClient(
    HttpClient httpClient,
    IOptions<MigoApiSettings> settings) : IMigoTipoCambioApiClient
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public async Task<MigoTipoCambioItemDto?> ObtenerPorFechaAsync(DateOnly fecha, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        ValidarConfiguracion(cfg);

        using var request = CrearRequest(
            cfg.ExchangeDatePath,
            new Dictionary<string, string>
            {
                ["token"] = cfg.Token.Trim(),
                ["fecha"] = fecha.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            });

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"La API de Migo devolvio HTTP {(int)response.StatusCode} al consultar el tipo de cambio por fecha.");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<MigoTipoCambioFechaResponse>(stream, JsonOptions, cancellationToken);
        if (payload is null || !payload.Success || string.IsNullOrWhiteSpace(payload.Fecha))
        {
            return null;
        }

        return MapItem(payload.Fecha, payload.Moneda, payload.PrecioCompra, payload.PrecioVenta);
    }

    public async Task<IReadOnlyCollection<MigoTipoCambioItemDto>> ObtenerPorRangoAsync(DateOnly fechaInicio, DateOnly fechaFin, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        ValidarConfiguracion(cfg);

        using var request = CrearRequest(
            cfg.ExchangeRangePath,
            new Dictionary<string, string>
            {
                ["token"] = cfg.Token.Trim(),
                ["fecha_inicio"] = fechaInicio.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ["fecha_fin"] = fechaFin.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            });

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"La API de Migo devolvio HTTP {(int)response.StatusCode} al consultar el tipo de cambio por rango.");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<MigoTipoCambioRangoResponse>(stream, JsonOptions, cancellationToken);
        if (payload is null || !payload.Success || payload.Data is null || payload.Data.Count == 0)
        {
            return Array.Empty<MigoTipoCambioItemDto>();
        }

        return payload.Data
            .Select(x => MapItem(x.Fecha, x.Moneda, x.PrecioCompra, x.PrecioVenta))
            .Where(x => x is not null)
            .Cast<MigoTipoCambioItemDto>()
            .ToList();
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

    private static MigoTipoCambioItemDto? MapItem(string? fecha, string? moneda, string? precioCompra, string? precioVenta)
    {
        if (!DateOnly.TryParseExact(fecha, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var fechaValue))
        {
            return null;
        }

        if (!decimal.TryParse(precioCompra, NumberStyles.Number, CultureInfo.InvariantCulture, out var compra))
        {
            return null;
        }

        if (!decimal.TryParse(precioVenta, NumberStyles.Number, CultureInfo.InvariantCulture, out var venta))
        {
            return null;
        }

        return new MigoTipoCambioItemDto
        {
            Fecha = fechaValue,
            Moneda = (moneda ?? string.Empty).Trim().ToUpperInvariant(),
            PrecioCompra = compra,
            PrecioVenta = venta
        };
    }

    private sealed class MigoTipoCambioFechaResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("fecha")]
        public string? Fecha { get; set; }

        [JsonPropertyName("moneda")]
        public string? Moneda { get; set; }

        [JsonPropertyName("precio_compra")]
        public string? PrecioCompra { get; set; }

        [JsonPropertyName("precio_venta")]
        public string? PrecioVenta { get; set; }
    }

    private sealed class MigoTipoCambioRangoResponse
    {
        [JsonPropertyName("success")]
        public bool Success { get; set; }

        [JsonPropertyName("data")]
        public List<MigoTipoCambioFechaResponse> Data { get; set; } = [];
    }
}
