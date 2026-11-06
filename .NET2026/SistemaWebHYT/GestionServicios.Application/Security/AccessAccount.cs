namespace GestionServicios.Application.Security;

public sealed record AccessAccount(
    string Id, string UserName, string DisplayName, string? PasswordHash,
    string SecurityStamp, bool Active, DateTimeOffset? LockoutEnd,
    IReadOnlyList<string> Roles, IReadOnlyList<string> Permissions);

public sealed record AccessResult(bool Succeeded, AccessAccount? Account);

// UsuarioId siempre procede de la identidad autenticada o del resultado de búsqueda,
// nunca de un campo editable del formulario de acceso.
public interface IAccessRepository
{
    Task<AccessAccount?> FindByNameAsync(string normalizedName, CancellationToken cancellationToken = default);
    Task<AccessAccount?> FindByIdAsync(string id, CancellationToken cancellationToken = default);
    Task<bool> CompleteAttemptAsync(string id, string expectedStamp, bool passwordValid,
        string? upgradedHash, Guid correlationId, CancellationToken cancellationToken = default);
    Task RecordLogoutAsync(string id, Guid correlationId, CancellationToken cancellationToken = default);
}
