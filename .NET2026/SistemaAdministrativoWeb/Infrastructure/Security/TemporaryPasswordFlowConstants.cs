namespace SistemaAdministrativoWeb.Infrastructure.Security;

public static class TemporaryPasswordFlowConstants
{
    public const string RequirePasswordChangeClaimType = "fralsecont:require-password-change";
    public const string RequirePasswordChangeClaimValue = "1";
    public const string VerificationSessionKey = "Auth:TemporaryPassword:VerifiedUserId";
}
