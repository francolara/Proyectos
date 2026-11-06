// Firma: FRANCO LARA - 05/11/2026 | Catálogos dependientes de ubigeo.
using System.Data;
using System.Text.Json;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Maintenance;

public sealed class SqlMaintenanceRepository(SqlConnectionFactory connections) : IMaintenanceRepository
{
    private static string Procedure(EntityDefinition entity, string operation)
    {
        if (!ReferenceEquals(MaintenanceDefinitions.Find(entity.Key), entity))
            throw new ArgumentException("Mantenimiento no reconocido.", nameof(entity));
        return $"dbo.spu_Mant{entity.DatabaseName}_{operation}";
    }

    private async Task<T> Execute<T>(string procedure, string actor, Action<SqlParameterCollection> parameters,
        Func<SqlCommand, Task<T>> execute, CancellationToken ct)
    {
        try
        {
            await using var connection = connections.CreateConnection();
            await connection.OpenAsync(ct);
            await using var command = new SqlCommand(procedure, connection) { CommandType = CommandType.StoredProcedure };
            command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
            parameters(command.Parameters);
            return await execute(command);
        }
        catch (SqlException error) when (error.Message.Contains("[HYT_", StringComparison.Ordinal))
        {
            var start = error.Message.IndexOf(']');
            throw new MaintenanceException(start >= 0 ? error.Message[(start + 1)..].Trim() : "No se pudo guardar el registro.");
        }
        catch (SqlException error) when (error.Number is 2601 or 2627)
        { throw new MaintenanceException("Ya existe un registro con ese código. Revisa los datos."); }
    }

    private static string? Text(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
    private static DateTimeOffset Utc(SqlDataReader reader, int ordinal) => new(DateTime.SpecifyKind(reader.GetDateTime(ordinal), DateTimeKind.Utc));

    public Task<MaterialTariffPage> MaterialTariffsAsync(string actorId,long materialId,int page,CancellationToken ct=default,string? search=null)
        =>Execute("dbo.spu_MaterialTarifas_Listar",actorId,p=>{p.Add("@MaterialId",SqlDbType.BigInt).Value=materialId;p.Add("@Pagina",SqlDbType.Int).Value=page;p.Add("@Busqueda",SqlDbType.NVarChar,200).Value=(object?)search??DBNull.Value;},async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);await reader.ReadAsync(ct);var rows=JsonSerializer.Deserialize<JsonElement[]>(reader.GetString(0))??[];
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new MaterialTariffPage(rows,reader.GetInt64(0),page);
        },ct);
    public Task<long> SaveMaterialTariffAsync(string actor,long materialId,long? id,byte[]? version,string tax,decimal? pen,decimal? usd,bool active,CancellationToken ct=default)
        =>Execute("dbo.spu_MaterialTarifa_Guardar",actor,p=>
        {
            p.Add("@MaterialId",SqlDbType.BigInt).Value=materialId;p.Add("@Id",SqlDbType.BigInt).Value=(object?)id??DBNull.Value;
            p.Add("@VersionOriginal",SqlDbType.Binary,8).Value=(object?)version??DBNull.Value;p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();
            p.Add("@CodigoTipoIGV",SqlDbType.NVarChar,10).Value=tax;
            var penParameter=p.Add("@PrecioPEN",SqlDbType.Decimal);penParameter.Precision=19;penParameter.Scale=6;penParameter.Value=(object?)pen??DBNull.Value;
            var usdParameter=p.Add("@PrecioUSD",SqlDbType.Decimal);usdParameter.Precision=19;usdParameter.Scale=6;usdParameter.Value=(object?)usd??DBNull.Value;
            p.Add("@Activo",SqlDbType.Bit).Value=active;
        },async command=>Convert.ToInt64(await command.ExecuteScalarAsync(ct)),ct);

    public Task<MaintenancePage> ListAsync(EntityDefinition entity, string actorId, string? search, int page,
        long? parentId = null, bool activeOnly = false, CancellationToken cancellationToken = default)
        => Execute(Procedure(entity, "Listar"), actorId, p =>
        {
            p.Add("@Pagina", SqlDbType.Int).Value = page;
            p.Add("@Busqueda", SqlDbType.NVarChar, 200).Value = (object?)search ?? DBNull.Value;
            p.Add("@PadreId", SqlDbType.BigInt).Value = (object?)parentId ?? DBNull.Value;
            p.Add("@SoloActivos", SqlDbType.Bit).Value = activeOnly;
        }, async command =>
        {
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            var rows = new List<MaintenanceRow>();
            while (await reader.ReadAsync(cancellationToken))
                rows.Add(new(Convert.ToInt64(reader.GetValue(0)), reader.GetString(1), reader.GetString(2), reader.GetString(3), Text(reader, 4),
                    Utc(reader, 5), reader.GetString(6), reader.IsDBNull(7) ? null : Utc(reader, 7), Text(reader, 8), entity.Key is "tiendas" or "contactos" ? Text(reader,9) : null, entity.Key is "tiendas" or "contactos" ? Text(reader,10) : null));
            await reader.NextResultAsync(cancellationToken);
            var total = await reader.ReadAsync(cancellationToken) ? reader.GetInt64(0) : 0;
            return new MaintenancePage(rows, total, page);
        }, cancellationToken);

    public Task<RecordSnapshot?> GetAsync(EntityDefinition entity, string actorId, long id, int auditPage = 1,
        CancellationToken cancellationToken = default)
        => Execute<RecordSnapshot?>(Procedure(entity, "Obtener"), actorId, p =>
        {
            p.Add("@Id", SqlDbType.BigInt).Value = id;
            p.Add("@PaginaAuditoria", SqlDbType.Int).Value = auditPage;
        }, async command =>
        {
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            if (!await reader.ReadAsync(cancellationToken) || reader.IsDBNull(0)) return null;
            var json = reader.GetString(0);
            if (string.IsNullOrWhiteSpace(json)) return null;
            var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json)!;
            var zones = values.TryGetValue("Zonas", out var zoneJson) && zoneJson.ValueKind == JsonValueKind.Array
                ? zoneJson.EnumerateArray().Select(item => item.GetProperty("Id").GetInt32()).ToArray() : [];
            var audit = new List<AuditEntry>();
            await reader.NextResultAsync(cancellationToken);
            while (await reader.ReadAsync(cancellationToken))
                audit.Add(new(reader.GetString(0), reader.GetString(1), Utc(reader, 2), Text(reader, 3), Text(reader, 4), Text(reader, 5)));
            await reader.NextResultAsync(cancellationToken);
            var total = await reader.ReadAsync(cancellationToken) ? reader.GetInt64(0) : 0;
            return new RecordSnapshot(id, values, zones, audit, total);
        }, cancellationToken);

    public Task<long> SaveAsync(EntityDefinition entity, string actorId, long? id, byte[]? version, string payload,
        CancellationToken cancellationToken = default)
        => Execute(Procedure(entity, "Guardar"), actorId, p =>
        {
            p.Add("@Id", SqlDbType.BigInt).Value = (object?)id ?? DBNull.Value;
            p.Add("@VersionOriginal", SqlDbType.Binary, 8).Value = (object?)version ?? DBNull.Value;
            p.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
            p.Add("@Datos", SqlDbType.NVarChar, -1).Value = payload;
        }, async command => Convert.ToInt64(await command.ExecuteScalarAsync(cancellationToken)), cancellationToken);

    public Task<IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>>> CatalogsAsync(string actorId, CancellationToken cancellationToken = default)
        => Execute<IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>>>("dbo.spu_MantCatalogos", actorId, _ => { }, async command =>
        {
            var result = new Dictionary<string, IReadOnlyList<CatalogOption>>();
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);
            foreach (var name in new[] { "Currency", "TaxTreatment", "Departments", "Provinces", "Districts" })
            {
                var options = new List<CatalogOption>();
                while (await reader.ReadAsync(cancellationToken)) options.Add(new(reader.GetString(0), reader.GetString(1)));
                result.Add(name, options);
                await reader.NextResultAsync(cancellationToken);
            }
            result.Add("EntityType", [new("JURIDICA", "Persona jurídica"), new("NATURAL", "Persona natural")]);
            return result;
        }, cancellationToken);

    public async Task PublishVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default)
        => _ = await VersionOperation("Publicar", actorId, versionId, expectedVersion, cancellationToken);
    public Task<long> CloneVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default)
        => VersionOperation("Clonar", actorId, versionId, expectedVersion, cancellationToken);
    private Task<long> VersionOperation(string operation, string actor, long id, byte[] version, CancellationToken ct)
        => Execute($"dbo.spu_MantTarifarioVersion_{operation}", actor, p =>
        {
            p.Add("@Id", SqlDbType.BigInt).Value = id;
            p.Add("@VersionOriginal", SqlDbType.Binary, 8).Value = version;
            p.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
        }, async command => Convert.ToInt64(await command.ExecuteScalarAsync(ct)), ct);
}
