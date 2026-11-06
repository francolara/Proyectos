using System.Security.Claims;
using GestionServicios.Application.Security;
using Microsoft.AspNetCore.Identity;

namespace GestionServicios.Web.Security;

public sealed class AccessService(IAccessRepository repository, IPasswordHasher<AccessAccount> hasher, TimeProvider clock)
{
    public const string StampClaim = "hyt:stamp";
    public const string PermissionClaim = "hyt:permission";
    private static readonly AccessAccount Dummy = new("", "", "", null, "", false, null, [], []);
    private readonly string dummyHash = hasher.HashPassword(Dummy, Guid.NewGuid().ToString("N"));

    public async Task<AccessResult> SignInAsync(string userName, string password, CancellationToken cancellationToken)
    {
        var account = await repository.FindByNameAsync(userName.Trim().ToUpperInvariant(), cancellationToken);
        var result = PasswordVerificationResult.Failed;
        try
        {
            result = hasher.VerifyHashedPassword(account ?? Dummy, account?.PasswordHash ?? dummyHash, password);
        }
        catch (FormatException) { /* Hash inválido: rechazar sin mostrar detalles. */ }
        if (account is null) return new(false, null);
        var valid = result != PasswordVerificationResult.Failed && account.Active &&
            !string.IsNullOrWhiteSpace(account.SecurityStamp) &&
            (account.LockoutEnd is null || account.LockoutEnd <= clock.GetUtcNow());
        var upgrade = valid && result == PasswordVerificationResult.SuccessRehashNeeded ? hasher.HashPassword(account, password) : null;
        var completed = await repository.CompleteAttemptAsync(account.Id, account.SecurityStamp, valid, upgrade, Guid.NewGuid(), cancellationToken);
        return new(completed && valid, completed && valid ? account : null);
    }

    public static ClaimsPrincipal Principal(AccessAccount account, long expiresUtcTicks)
    {
        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, account.Id), new(ClaimTypes.Name, account.DisplayName),
            new(StampClaim, account.SecurityStamp),
            new("hyt:expires", expiresUtcTicks.ToString(System.Globalization.CultureInfo.InvariantCulture))
        };
        claims.AddRange(account.Roles.Select(role => new Claim(ClaimTypes.Role, role)));
        claims.AddRange(account.Permissions.Select(permission => new Claim(PermissionClaim, permission)));
        return new(new ClaimsIdentity(claims, "HytCookie"));
    }
}
