using System.Text.Json;
namespace GestionServicios.Application.Security;
public sealed record SecurityPage(IReadOnlyList<JsonElement> Rows,long Total,int Page);
public sealed record SecurityAuditPage(SecurityPage Data, IReadOnlyList<JsonElement> Options);
public sealed record SecurityUserCommand(string Id,string? Version,string Names,string Surnames,string Email,string RoleId,bool Active,string? PasswordHash);
public sealed record SecurityRoleCommand(string Id,string? Version,string Name,IReadOnlyList<string> Permissions);
public interface ISecurityAdministrationRepository
{
    Task<SecurityPage> UsersAsync(string actor,string? search,bool? active,string? situation,int page,CancellationToken ct);
    Task<JsonElement?> UserAsync(string actor,string id,CancellationToken ct);
    Task<(IReadOnlyList<JsonElement> Roles,IReadOnlyList<JsonElement> Permissions)> CatalogsAsync(string actor,CancellationToken ct);
    Task SaveUserAsync(string actor,SecurityUserCommand input,CancellationToken ct);
    Task UserActionAsync(string actor,string id,string version,string action,string? hash,string reason,CancellationToken ct);
    Task<SecurityPage> RolesAsync(string actor,string? search,int page,CancellationToken ct);
    Task<JsonElement?> RoleAsync(string actor,string id,CancellationToken ct);
    Task SaveRoleAsync(string actor,SecurityRoleCommand input,CancellationToken ct);
    Task ChangePasswordAsync(string actor,string expectedStamp,string expectedHash,string newHash,bool mandatory,CancellationToken ct);
    Task<SecurityAuditPage> AuditAsync(string actor,DateTimeOffset from,DateTimeOffset until,string? search,string? user,string? entity,string? action,int page,CancellationToken ct);
    Task<JsonElement?> AuditEventAsync(string actor,long id,CancellationToken ct);
}
