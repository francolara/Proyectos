using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Caching.Memory;

namespace GestionServicios.Web.Security;

// Firma: FRANCO LARA - 05/11/2026 | Captcha de respaldo ligado al navegador, cinco minutos y un solo intento; almacenamiento acotado.
public sealed class ManualCaptchaService(TimeProvider clock) : IDisposable
{
    private const string CookieName = "__Host-Hyt.Captcha";
    private const string Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    private readonly MemoryCache cache = new(new MemoryCacheOptions { SizeLimit = 5000 });
    private readonly object gate = new();
    private sealed record Challenge(string Browser, string Code, DateTimeOffset Expires);

    public Guid Create(HttpContext context)
    {
        var browser = context.Request.Cookies[CookieName];
        if (!Guid.TryParseExact(browser, "N", out _))
        {
            browser = Guid.NewGuid().ToString("N");
            context.Response.Cookies.Append(CookieName, browser, new CookieOptions
            {
                Secure = true, HttpOnly = true, SameSite = SameSiteMode.Lax, Path = "/",
                IsEssential = true, MaxAge = TimeSpan.FromMinutes(30)
            });
        }
        var code = new string(Enumerable.Range(0, 6).Select(_ => Alphabet[RandomNumberGenerator.GetInt32(Alphabet.Length)]).ToArray());
        var id = Guid.NewGuid();
        cache.Set(id, new Challenge(browser!, code, clock.GetUtcNow().AddMinutes(5)), new MemoryCacheEntryOptions { Size = 1, AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(5) });
        return id;
    }

    private Challenge? Find(HttpContext context, Guid id)
        => cache.TryGetValue(id, out Challenge? challenge) && challenge is not null &&
           challenge.Expires > clock.GetUtcNow() && challenge.Browser == context.Request.Cookies[CookieName] ? challenge : null;

    public bool Validate(HttpContext context, Guid? id, string? answer)
    {
        if (id is null) return false;
        lock (gate)
        {
            var challenge = Find(context, id.Value);
            if (challenge is null) return false;
            cache.Remove(id.Value); // Un intento incluso cuando la respuesta es incorrecta.
            var supplied = answer?.Trim().ToUpperInvariant();
            return supplied is { Length: 6 } && CryptographicOperations.FixedTimeEquals(
                Encoding.ASCII.GetBytes(challenge.Code), Encoding.ASCII.GetBytes(supplied));
        }
    }

    public string? Image(HttpContext context, Guid id)
    {
        var challenge = Find(context, id);
        if (challenge is null) return null;
        var svg = new StringBuilder("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"80\" viewBox=\"0 0 260 80\"><rect width=\"260\" height=\"80\" rx=\"6\" fill=\"#f3f3f3\"/>");
        for (var line = 0; line < 6; line++)
            svg.Append($"<path d=\"M 0 {RandomNumberGenerator.GetInt32(8, 73)} L 260 {RandomNumberGenerator.GetInt32(8, 73)}\" stroke=\"#c9c9c9\" fill=\"none\"/>");
        for (var index = 0; index < challenge.Code.Length; index++)
        {
            var x = 25 + index * 37;
            var y = RandomNumberGenerator.GetInt32(48, 58);
            svg.Append($"<text x=\"{x}\" y=\"{y}\" transform=\"rotate({RandomNumberGenerator.GetInt32(-12, 13)} {x} {y})\" fill=\"#181818\" font-family=\"monospace\" font-weight=\"bold\" font-size=\"34\">{challenge.Code[index]}</text>");
        }
        return svg.Append("</svg>").ToString();
    }

    public void Dispose() => cache.Dispose();
}
