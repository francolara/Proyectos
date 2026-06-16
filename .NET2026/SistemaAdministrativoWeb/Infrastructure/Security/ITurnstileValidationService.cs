using SistemaAdministrativoWeb.Configuration;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public interface ITurnstileValidationService
{
    Task<TurnstileVerifyResponse> VerifyAsync(string token, string? remoteIp, CancellationToken cancellationToken = default);
}
