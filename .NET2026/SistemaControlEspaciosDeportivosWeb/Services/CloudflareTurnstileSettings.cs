using System.Text.Json.Serialization;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public sealed class CloudflareTurnstileSettings
{
    public string SiteKey { get; set; } = string.Empty;
    public string SecretKey { get; set; } = string.Empty;
    public string VerifyUrl { get; set; } = "https://challenges.cloudflare.com/turnstile/v0/siteverify";
    public int LoginFailuresBeforeChallenge { get; set; } = 2;
    public int ResendAttemptsBeforeChallenge { get; set; } = 2;
    public int ForgotPasswordAttemptsBeforeChallenge { get; set; } = 2;
}

public sealed class TurnstileVerifyResponse
{
    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("error-codes")]
    public string[] ErrorCodes { get; set; } = Array.Empty<string>();
}
