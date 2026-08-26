using System.Text.Json.Serialization;

namespace SistemaAdministrativoWeb.Configuration;

public sealed class CloudflareTurnstileSettings
{
    public const string SectionName = "FRALSECONT_CloudflareTurnstile";

    public string SiteKey { get; set; } = string.Empty;
    public string SecretKey { get; set; } = string.Empty;
    public string VerifyUrl { get; set; } = "https://challenges.cloudflare.com/turnstile/v0/siteverify";
    public int LoginFailuresBeforeChallenge { get; set; } = 2;
    public int ResendAttemptsBeforeChallenge { get; set; } = 2;
}

public sealed class TurnstileVerifyResponse
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("error-codes")]
    public string[] ErrorCodes { get; set; } = Array.Empty<string>();
}
