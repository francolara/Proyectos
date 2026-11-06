using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Options;
namespace GestionServicios.Web.Security;
public sealed class CloudflareTurnstileSettings
{
    public bool Enabled {get;set;}=true;
    public string SiteKey {get;set;}="";
    public string SecretKey {get;set;}="";
    public string[] AllowedHostnames {get;set;}=[];
}
public interface ITurnstileVerifier
{
    bool Required {get;}
    string SiteKey {get;}
    Task<bool> VerifyAsync(string? token,CancellationToken ct);
}
public sealed class CloudflareTurnstileVerifier(HttpClient client,IOptions<CloudflareTurnstileSettings> options,IWebHostEnvironment environment,TimeProvider clock):ITurnstileVerifier
{
    public bool Required=>!environment.IsDevelopment()||options.Value.Enabled;
    public string SiteKey=>options.Value.SiteKey;
    public async Task<bool> VerifyAsync(string? token,CancellationToken ct)
    {
        if(!Required)return true;
        var settings=options.Value;
        if(string.IsNullOrWhiteSpace(settings.SiteKey)||string.IsNullOrWhiteSpace(settings.SecretKey)||settings.AllowedHostnames.Length==0||string.IsNullOrWhiteSpace(token)||token.Length>2048)return false;
        try
        {
            using var response=await client.PostAsync("https://challenges.cloudflare.com/turnstile/v0/siteverify",new FormUrlEncodedContent(new Dictionary<string,string>{["secret"]=settings.SecretKey,["response"]=token,["idempotency_key"]=Guid.NewGuid().ToString()}),ct);
            if(!response.IsSuccessStatusCode)return false;
            var result=await response.Content.ReadFromJsonAsync<TurnstileResult>(cancellationToken:ct);
            return result is {Success:true,Action:"login"}&&settings.AllowedHostnames.Any(host=>string.Equals(host,result.Hostname,StringComparison.OrdinalIgnoreCase))
                &&result.ChallengeTime.HasValue&&clock.GetUtcNow()-result.ChallengeTime.Value is var age&&age>=TimeSpan.FromMinutes(-1)&&age<=TimeSpan.FromMinutes(5);
        }
        catch(Exception error)when(error is HttpRequestException or JsonException||error is OperationCanceledException&&!ct.IsCancellationRequested){return false;}
    }
    private sealed class TurnstileResult
    {
        [JsonPropertyName("success")]public bool Success {get;set;}
        [JsonPropertyName("action")]public string? Action {get;set;}
        [JsonPropertyName("hostname")]public string? Hostname {get;set;}
        [JsonPropertyName("challenge_ts")]public DateTimeOffset? ChallengeTime {get;set;}
    }
}