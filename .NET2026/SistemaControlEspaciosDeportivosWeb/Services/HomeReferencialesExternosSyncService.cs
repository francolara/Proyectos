using System.Data;
using System.Globalization;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class HomeReferencialesExternosSyncService(
    IConfiguration configuration,
    IHttpClientFactory httpClientFactory,
    ISportCenterStoredProcedureService spService,
    ILogger<HomeReferencialesExternosSyncService> logger) : IHomeReferencialesExternosSyncService
{
    private readonly string _connectionString = configuration.GetConnectionString("DefaultConnection")
                                               ?? throw new InvalidOperationException("No se encontro DefaultConnection.");

    public async Task<ReferencialesExternosSyncResultadoViewModel> EjecutarBarridoAsync(
        string codigoUbigeo,
        int tipoDeporteSuperId,
        string palabraClave,
        int maxResultados,
        bool descargarTelefonos,
        bool descargarFotos,
        string usuario,
        CancellationToken cancellationToken = default)
    {
        var resultado = new ReferencialesExternosSyncResultadoViewModel();
        logger.LogInformation(
            "Inicio barrido referenciales externos. Ubigeo={CodigoUbigeo}, TipoDeporteSuperId={TipoDeporteSuperId}, MaxResultados={MaxResultados}, DescargarTelefonos={DescargarTelefonos}, DescargarFotos={DescargarFotos}.",
            codigoUbigeo,
            tipoDeporteSuperId,
            maxResultados,
            descargarTelefonos,
            descargarFotos);
        var apiKey = (configuration["GoogleMaps:ApiKeyServer"] ?? configuration["GoogleMaps:ApiKey"] ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(apiKey))
            throw new InvalidOperationException("No se encontro GoogleMaps:ApiKeyServer (ni fallback GoogleMaps:ApiKey) para ejecutar el barrido.");

        var ubigeo = await spService.UbigeoObtenerPorCodigoAsync(codigoUbigeo);
        if (ubigeo is null)
            throw new InvalidOperationException("No se encontro el ubigeo seleccionado.");

        var queryBase = $"{palabraClave} {ubigeo.Distrito}, {ubigeo.Provincia}, {ubigeo.Departamento}, Peru".Trim();
        var client = httpClientFactory.CreateClient("GooglePlacesTextSearch");
        string? nextPageToken = null;
        var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        while (resultado.Procesados < maxResultados)
        {
            if (!string.IsNullOrWhiteSpace(nextPageToken))
                await Task.Delay(2000, cancellationToken);

            var url = string.IsNullOrWhiteSpace(nextPageToken)
                ? $"maps/api/place/textsearch/json?query={Uri.EscapeDataString(queryBase)}&language=es&region=pe&key={Uri.EscapeDataString(apiKey)}"
                : $"maps/api/place/textsearch/json?pagetoken={Uri.EscapeDataString(nextPageToken)}&key={Uri.EscapeDataString(apiKey)}";

            using var response = await client.GetAsync(url, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                resultado.Errores++;
                resultado.Detalles.Add($"Google Places devolvio {(int)response.StatusCode}.");
                break;
            }

            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            var root = document.RootElement;
            var status = root.TryGetProperty("status", out var statusProp) ? statusProp.GetString() ?? string.Empty : string.Empty;
            var errorMessage = root.TryGetProperty("error_message", out var errorProp) ? errorProp.GetString() ?? string.Empty : string.Empty;
            if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(status, "ZERO_RESULTS", StringComparison.OrdinalIgnoreCase))
            {
                resultado.Errores++;
                resultado.Detalles.Add(string.IsNullOrWhiteSpace(errorMessage)
                    ? $"Google Places status: {status}."
                    : $"Google Places status: {status}. Detalle: {errorMessage}");
                break;
            }

            if (!root.TryGetProperty("results", out var places) || places.ValueKind != JsonValueKind.Array || places.GetArrayLength() == 0)
                break;

            foreach (var place in places.EnumerateArray())
            {
                if (resultado.Procesados >= maxResultados) break;

                var placeId = TryGetString(place, "place_id");
                if (string.IsNullOrWhiteSpace(placeId) || !vistos.Add(placeId))
                {
                    resultado.Omitidos++;
                    continue;
                }

                var name = TryGetString(place, "name");
                if (string.IsNullOrWhiteSpace(name))
                {
                    resultado.Omitidos++;
                    continue;
                }

                var direccion = TryGetString(place, "formatted_address");
                var mapsUrl = string.IsNullOrWhiteSpace(placeId) ? null : $"https://www.google.com/maps/place/?q=place_id:{placeId}";
                var detalle = await ObtenerDetallePlaceAsync(client, placeId!, apiKey, descargarTelefonos, descargarFotos, cancellationToken);
                decimal? lat = null;
                decimal? lng = null;
                if (place.TryGetProperty("geometry", out var geometry) &&
                    geometry.TryGetProperty("location", out var location))
                {
                    lat = TryGetDecimal(location, "lat");
                    lng = TryGetDecimal(location, "lng");
                }

                var upsert = await UpsertReferencialAsync(
                    placeId!,
                    name!,
                    codigoUbigeo,
                    tipoDeporteSuperId,
                    direccion,
                    mapsUrl,
                    detalle.TelefonoContacto,
                    detalle.FotoPrincipalUrl,
                    detalle.FotosUrlsCsv,
                    lat,
                    lng,
                    usuario,
                    cancellationToken);

                resultado.Procesados++;
                if (upsert == "INSERTADO") resultado.Insertados++;
                else if (upsert == "ACTUALIZADO") resultado.Actualizados++;
                else resultado.Omitidos++;
            }

            nextPageToken = root.TryGetProperty("next_page_token", out var nextTokenProp)
                ? (nextTokenProp.GetString() ?? string.Empty).Trim()
                : string.Empty;

            if (string.IsNullOrWhiteSpace(nextPageToken))
                break;
        }

        if (resultado.Procesados == 0 && resultado.Errores == 0)
            resultado.Detalles.Add("No se encontraron resultados para los criterios seleccionados.");

        logger.LogInformation(
            "Fin barrido referenciales externos. Procesados={Procesados}, Insertados={Insertados}, Actualizados={Actualizados}, Omitidos={Omitidos}, Errores={Errores}.",
            resultado.Procesados,
            resultado.Insertados,
            resultado.Actualizados,
            resultado.Omitidos,
            resultado.Errores);

        return resultado;
    }

    private async Task<string> UpsertReferencialAsync(
        string googlePlaceId,
        string nombreComplejo,
        string codigoUbigeo,
        int tipoDeporteSuperId,
        string? direccion,
        string? googleMapsUrl,
        string? telefonoContacto,
        string? fotoPrincipalUrl,
        string? fotosUrlsCsv,
        decimal? latitud,
        decimal? longitud,
        string usuario,
        CancellationToken cancellationToken)
    {
        await using var cn = new SqlConnection(_connectionString);
        await cn.OpenAsync(cancellationToken);
        await using var cmd = new SqlCommand("Sp_Home_ReferencialExterno_UpsertDesdeGoogle", cn)
        {
            CommandType = CommandType.StoredProcedure
        };

        cmd.Parameters.Add("@GooglePlaceId", SqlDbType.NVarChar, 200).Value = googlePlaceId;
        cmd.Parameters.Add("@NombreComplejo", SqlDbType.NVarChar, 180).Value = nombreComplejo;
        cmd.Parameters.Add("@NombreEspacio", SqlDbType.NVarChar, 150).Value = DBNull.Value;
        cmd.Parameters.Add("@CodigoReferencia", SqlDbType.NVarChar, 50).Value = $"G-{googlePlaceId[..Math.Min(12, googlePlaceId.Length)]}";
        cmd.Parameters.Add("@CodigoUbigeo", SqlDbType.Char, 6).Value = codigoUbigeo;
        cmd.Parameters.Add("@TipoDeporteSuperId", SqlDbType.Int).Value = tipoDeporteSuperId;
        cmd.Parameters.Add("@Direccion", SqlDbType.NVarChar, 250).Value = string.IsNullOrWhiteSpace(direccion) ? DBNull.Value : direccion.Trim();
        cmd.Parameters.Add("@Referencia", SqlDbType.NVarChar, 1000).Value = DBNull.Value;
        cmd.Parameters.Add("@TelefonoContacto", SqlDbType.NVarChar, 40).Value = string.IsNullOrWhiteSpace(telefonoContacto) ? DBNull.Value : telefonoContacto.Trim();
        cmd.Parameters.Add("@CorreoContacto", SqlDbType.NVarChar, 200).Value = DBNull.Value;
        cmd.Parameters.Add("@WhatsappContacto", SqlDbType.NVarChar, 30).Value = DBNull.Value;
        cmd.Parameters.Add("@PermiteChatWhatsapp", SqlDbType.Bit).Value = false;
        var tarifaParam = cmd.Parameters.Add("@TarifaReferencial", SqlDbType.Decimal);
        tarifaParam.Precision = 10;
        tarifaParam.Scale = 2;
        tarifaParam.Value = DBNull.Value;
        cmd.Parameters.Add("@TieneIluminacion", SqlDbType.Bit).Value = false;
        cmd.Parameters.Add("@Techada", SqlDbType.Bit).Value = false;
        cmd.Parameters.Add("@GoogleMapsUrl", SqlDbType.NVarChar, 500).Value = string.IsNullOrWhiteSpace(googleMapsUrl) ? DBNull.Value : googleMapsUrl.Trim();
        cmd.Parameters.Add("@FotoPrincipalUrl", SqlDbType.NVarChar, 500).Value = string.IsNullOrWhiteSpace(fotoPrincipalUrl) ? DBNull.Value : fotoPrincipalUrl.Trim();
        cmd.Parameters.Add("@FotosUrlsCsv", SqlDbType.NVarChar).Value = string.IsNullOrWhiteSpace(fotosUrlsCsv) ? DBNull.Value : fotosUrlsCsv.Trim();
        var latParam = cmd.Parameters.Add("@LatitudReferencia", SqlDbType.Decimal);
        latParam.Precision = 10;
        latParam.Scale = 7;
        latParam.Value = latitud.HasValue ? latitud.Value : DBNull.Value;

        var lngParam = cmd.Parameters.Add("@LongitudReferencia", SqlDbType.Decimal);
        lngParam.Precision = 10;
        lngParam.Scale = 7;
        lngParam.Value = longitud.HasValue ? longitud.Value : DBNull.Value;
        cmd.Parameters.Add("@Activo", SqlDbType.Bit).Value = true;
        cmd.Parameters.Add("@Usuario", SqlDbType.NVarChar, 200).Value = string.IsNullOrWhiteSpace(usuario) ? "owner-platform" : usuario.Trim();

        var accion = await cmd.ExecuteScalarAsync(cancellationToken);
        var valor = Convert.ToString(accion, CultureInfo.InvariantCulture) ?? string.Empty;
        return string.IsNullOrWhiteSpace(valor) ? "OMITIDO" : valor.Trim().ToUpperInvariant();
    }

    private async Task<GooglePlaceDetalle> ObtenerDetallePlaceAsync(
        HttpClient client,
        string placeId,
        string apiKey,
        bool descargarTelefonos,
        bool descargarFotos,
        CancellationToken cancellationToken)
    {
        if (!descargarTelefonos && !descargarFotos)
            return GooglePlaceDetalle.Empty;

        try
        {
            var fields = new List<string>();
            if (descargarTelefonos)
            {
                fields.Add("formatted_phone_number");
                fields.Add("international_phone_number");
            }
            if (descargarFotos)
                fields.Add("photos");

            if (fields.Count == 0)
                return GooglePlaceDetalle.Empty;

            var fieldsCsv = string.Join(",", fields);
            var url = $"maps/api/place/details/json?place_id={Uri.EscapeDataString(placeId)}&fields={Uri.EscapeDataString(fieldsCsv)}&language=es&region=pe&key={Uri.EscapeDataString(apiKey)}";
            using var response = await client.GetAsync(url, cancellationToken);
            if (!response.IsSuccessStatusCode)
                return GooglePlaceDetalle.Empty;

            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            var root = document.RootElement;
            var status = root.TryGetProperty("status", out var statusProp) ? statusProp.GetString() ?? string.Empty : string.Empty;
            if (!string.Equals(status, "OK", StringComparison.OrdinalIgnoreCase))
                return GooglePlaceDetalle.Empty;

            if (!root.TryGetProperty("result", out var result) || result.ValueKind != JsonValueKind.Object)
                return GooglePlaceDetalle.Empty;

            var telefono = descargarTelefonos
                ? TryGetString(result, "international_phone_number") ?? TryGetString(result, "formatted_phone_number")
                : null;

            var fotos = new List<string>();
            if (descargarFotos && result.TryGetProperty("photos", out var photosNode) && photosNode.ValueKind == JsonValueKind.Array)
            {
                foreach (var foto in photosNode.EnumerateArray())
                {
                    var referenciaFoto = TryGetString(foto, "photo_reference");
                    if (string.IsNullOrWhiteSpace(referenciaFoto))
                        continue;

                    fotos.Add($"https://maps.googleapis.com/maps/api/place/photo?maxwidth=1280&photo_reference={Uri.EscapeDataString(referenciaFoto)}&key={Uri.EscapeDataString(apiKey)}");
                    if (fotos.Count >= 5) break;
                }
            }

            var fotoPrincipal = fotos.Count > 0 ? fotos[0] : null;
            var fotosCsv = fotos.Count > 0 ? string.Join(",", fotos) : null;
            return new GooglePlaceDetalle(telefono, fotoPrincipal, fotosCsv);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "No se pudo obtener Place Details para PlaceId={PlaceId}. Se continua con datos base.", placeId);
            return GooglePlaceDetalle.Empty;
        }
    }

    private static string? TryGetString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;

        return prop.ValueKind == JsonValueKind.String ? prop.GetString() : prop.ToString();
    }

    private static decimal? TryGetDecimal(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop) || prop.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;

        if (prop.ValueKind == JsonValueKind.Number && prop.TryGetDecimal(out var value))
            return value;

        if (decimal.TryParse(prop.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            return value;

        return null;
    }

    private readonly record struct GooglePlaceDetalle(
        string? TelefonoContacto,
        string? FotoPrincipalUrl,
        string? FotosUrlsCsv)
    {
        public static GooglePlaceDetalle Empty => new(null, null, null);
    }
}
