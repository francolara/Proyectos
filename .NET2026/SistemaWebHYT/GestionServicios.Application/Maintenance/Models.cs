using System.Text.Json;
namespace GestionServicios.Application.Maintenance;

public sealed record FieldDefinition(string Name, string Label, string Kind, bool Required, int MaxLength,
    string? Lookup = null, string? Default = null, string? Catalog = null);
public sealed record EntityDefinition(string Key, string Title, string DatabaseName, string Permission,
    IReadOnlyList<FieldDefinition> Fields, string? ParentField, bool HasZones);
public sealed record MaintenanceRow(long Id, string Code, string Description, string Status, string? Extra,
    DateTimeOffset CreatedAt, string CreatedBy, DateTimeOffset? UpdatedAt, string? UpdatedBy, string? ClientCode = null, string? ClientName = null, string? StoreName = null);
public sealed record MaintenancePage(IReadOnlyList<MaintenanceRow> Rows, long Total, int Page);
public sealed record AuditEntry(string Action, string User, DateTimeOffset RecordedAt, string? PreviousState, string? NewState, string? Changes);
public sealed record RecordSnapshot(long Id, IReadOnlyDictionary<string, JsonElement> Values,
    IReadOnlyList<int> Zones, IReadOnlyList<AuditEntry> Audit, long AuditCount)
{
    public string Text(string key) => Values.TryGetValue(key, out var value) && value.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined) ? value.ToString() : "";
}
public sealed record CatalogOption(string Code, string Label);
public sealed record ValidationIssue(string Field, string Message);
public sealed class MaintenanceException(string message) : Exception(message);

public sealed record MaterialTariffPage(IReadOnlyList<JsonElement> Rows,long Total,int Page);

public interface IMaintenanceRepository
{
    Task<long> SaveMaterialTariffAsync(string actor,long materialId,long? id,byte[]? version,string tax,decimal? pen,decimal? usd,bool active,CancellationToken ct=default);
    Task<MaterialTariffPage> MaterialTariffsAsync(string actorId,long materialId,int page,CancellationToken ct=default,string? search=null);
    Task<MaintenancePage> ListAsync(EntityDefinition entity, string actorId, string? search, int page,
        long? parentId = null, bool activeOnly = false, CancellationToken cancellationToken = default);
    Task<RecordSnapshot?> GetAsync(EntityDefinition entity, string actorId, long id, int auditPage = 1, CancellationToken cancellationToken = default);
    Task<long> SaveAsync(EntityDefinition entity, string actorId, long? id, byte[]? version, string payload, CancellationToken cancellationToken = default);
    Task<IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>>> CatalogsAsync(string actorId, CancellationToken cancellationToken = default);
    Task PublishVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default);
    Task<long> CloneVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default);
}

public sealed class ClientDocumentLink
{
    public long? Id { get; set; }
    public string? Description { get; set; }
    public string? Url { get; set; }
}
