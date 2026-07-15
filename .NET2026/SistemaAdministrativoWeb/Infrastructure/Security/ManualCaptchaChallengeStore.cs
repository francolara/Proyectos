using System.Security.Cryptography;

namespace SistemaAdministrativoWeb.Infrastructure.Security;

public static class ManualCaptchaChallengeStore
{
    private const string Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

    public static string GetOrCreate(HttpContext httpContext, string scope)
    {
        var sessionKey = BuildSessionKey(scope);
        var currentCode = httpContext.Session.GetString(sessionKey);
        if (!string.IsNullOrWhiteSpace(currentCode))
        {
            return currentCode;
        }

        return Refresh(httpContext, scope);
    }

    public static string Refresh(HttpContext httpContext, string scope)
    {
        var code = GenerateCode();
        httpContext.Session.SetString(BuildSessionKey(scope), code);
        return code;
    }

    public static void Clear(HttpContext httpContext, string scope)
    {
        httpContext.Session.Remove(BuildSessionKey(scope));
    }

    public static bool Validate(HttpContext httpContext, string scope, string? userInput)
    {
        var expectedCode = httpContext.Session.GetString(BuildSessionKey(scope));
        if (string.IsNullOrWhiteSpace(expectedCode))
        {
            return false;
        }

        return string.Equals(
            Normalize(userInput),
            Normalize(expectedCode),
            StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildSessionKey(string scope) => $"Auth:ManualCaptcha:{scope}";

    private static string Normalize(string? value)
        => (value ?? string.Empty).Trim().Replace(" ", string.Empty, StringComparison.Ordinal);

    private static string GenerateCode()
    {
        Span<byte> buffer = stackalloc byte[5];
        RandomNumberGenerator.Fill(buffer);
        return string.Create(buffer.Length, buffer, static (span, source) =>
        {
            for (var i = 0; i < source.Length; i++)
            {
                span[i] = Alphabet[source[i] % Alphabet.Length];
            }
        });
    }
}
