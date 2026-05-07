using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Options;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public sealed class TurnstileValidationService(HttpClient httpClient, IOptions<CloudflareTurnstileSettings> settings) : ITurnstileValidationService
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public async Task<TurnstileVerifyResponse> VerifyAsync(string token, string? remoteIp, CancellationToken cancellationToken = default)
    {
        var cfg = settings.Value;
        if (string.IsNullOrWhiteSpace(cfg.SecretKey) || string.IsNullOrWhiteSpace(cfg.VerifyUrl))
        {
            return new TurnstileVerifyResponse
            {
                Success = false,
                ErrorCodes = new[] { "turnstile-not-configured" }
            };
        }

        using var request = new HttpRequestMessage(HttpMethod.Post, cfg.VerifyUrl)
        {
            Content = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["secret"] = cfg.SecretKey,
                ["response"] = token,
                ["remoteip"] = remoteIp ?? string.Empty
            })
        };
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var response = await httpClient.SendAsync(request, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return new TurnstileVerifyResponse
            {
                Success = false,
                ErrorCodes = new[] { $"http-{(int)response.StatusCode}" }
            };
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var payload = await JsonSerializer.DeserializeAsync<TurnstileVerifyResponse>(stream, JsonOptions, cancellationToken);
        return payload ?? new TurnstileVerifyResponse { Success = false, ErrorCodes = new[] { "invalid-json" } };
    }
}
