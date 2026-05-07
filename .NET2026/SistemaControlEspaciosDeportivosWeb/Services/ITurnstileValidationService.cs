namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface ITurnstileValidationService
{
    Task<TurnstileVerifyResponse> VerifyAsync(string token, string? remoteIp, CancellationToken cancellationToken = default);
}
