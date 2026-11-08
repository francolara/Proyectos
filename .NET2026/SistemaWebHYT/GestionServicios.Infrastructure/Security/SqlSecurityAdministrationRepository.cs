using System.Data;
using System.Text.Json;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Security;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Security;

public sealed class SqlSecurityAdministrationRepository(SqlConnectionFactory connections) : ISecurityAdministrationRepository
{
    public Task<SecurityPage> UsersAsync(string actor, string? search, bool? active, string? situation, int page, CancellationToken ct) =>
        Query("dbo.spu_Usuarios_Listar", actor, parameters =>
        {
            Text(parameters, "@Busqueda", search, 200);
            parameters.Add("@Activo", SqlDbType.Bit).Value = (object?)active ?? DBNull.Value;
            Text(parameters, "@Situacion", situation, 20);
            parameters.Add("@Pagina", SqlDbType.Int).Value = page;
        }, command => Page(command, page, ct), ct);

    public Task<JsonElement?> UserAsync(string actor, string id, CancellationToken ct) =>
        Query("dbo.spu_Usuarios_Obtener", actor, parameters => Text(parameters, "@Id", id, 128), command => One(command, ct), ct);

    public Task<(IReadOnlyList<JsonElement> Roles, IReadOnlyList<JsonElement> Permissions)> CatalogsAsync(string actor, CancellationToken ct) =>
        Query<(IReadOnlyList<JsonElement>, IReadOnlyList<JsonElement>)>("dbo.spu_Seguridad_Catalogos", actor, _ => { }, async command =>
        {
            await using var reader = await command.ExecuteReaderAsync(ct);
            var roles = await ReadRows(reader, ct);
            await reader.NextResultAsync(ct);
            return (roles, await ReadRows(reader, ct));
        }, ct);

    public Task SaveUserAsync(string actor, SecurityUserCommand input, CancellationToken ct) =>
        Save("dbo.spu_Usuarios_Guardar", actor, parameters => Text(parameters, "@Datos", JsonSerializer.Serialize(input)), ct);

    public Task UserActionAsync(string actor, string id, string version, string action, string? hash, string reason, CancellationToken ct) =>
        Save("dbo.spu_Usuarios_Accion", actor, parameters =>
        {
            Text(parameters, "@Id", id, 128);
            Text(parameters, "@VersionOriginal", version);
            Text(parameters, "@Accion", action, 30);
            Text(parameters, "@PasswordHash", hash);
            Text(parameters, "@Motivo", reason, 1000);
        }, ct);

    public Task<SecurityPage> RolesAsync(string actor, string? search, int page, CancellationToken ct) =>
        Query("dbo.spu_Perfiles_Listar", actor, parameters =>
        {
            Text(parameters, "@Busqueda", search, 200);
            parameters.Add("@Pagina", SqlDbType.Int).Value = page;
        }, command => Page(command, page, ct), ct);

    public Task<JsonElement?> RoleAsync(string actor, string id, CancellationToken ct) =>
        Query("dbo.spu_Perfiles_Obtener", actor, parameters => Text(parameters, "@Id", id, 128), command => One(command, ct), ct);

    public Task SaveRoleAsync(string actor, SecurityRoleCommand input, CancellationToken ct) =>
        Save("dbo.spu_Perfiles_Guardar", actor, parameters => Text(parameters, "@Datos", JsonSerializer.Serialize(input)), ct);

    public Task ChangePasswordAsync(string actor, string expectedStamp, string expectedHash, string newHash, bool mandatory, CancellationToken ct) =>
        Save("dbo.spu_Usuario_CambiarClave", actor, parameters =>
        {
            Text(parameters, "@SelloEsperado", expectedStamp);
            Text(parameters, "@HashEsperado", expectedHash);
            Text(parameters, "@PasswordHash", newHash);
            parameters.Add("@Obligatorio", SqlDbType.Bit).Value = mandatory;
        }, ct);

    public Task<SecurityAuditPage> AuditAsync(string actor, DateTimeOffset from, DateTimeOffset until,
        string? search, string? user, string? entity, string? action, int page, CancellationToken ct) =>
        Query("dbo.spu_Auditoria_Listar", actor, parameters =>
        {
            parameters.Add("@Desde", SqlDbType.DateTime2).Value = from.UtcDateTime;
            parameters.Add("@HastaExclusiva", SqlDbType.DateTime2).Value = until.UtcDateTime;
            Text(parameters, "@Busqueda", search, 200);
            Text(parameters, "@ActorFiltro", user, 128);
            Text(parameters, "@Entidad", entity, 128);
            Text(parameters, "@Accion", action, 60);
            parameters.Add("@Pagina", SqlDbType.Int).Value = page;
        }, async command =>
        {
            await using var reader = await command.ExecuteReaderAsync(ct);
            var data = await ReadPage(reader, page, ct);
            await reader.NextResultAsync(ct);
            return new SecurityAuditPage(data, await ReadRows(reader, ct));
        }, ct);

    public Task<JsonElement?> AuditEventAsync(string actor, long id, CancellationToken ct) =>
        Query("dbo.spu_Auditoria_Obtener", actor, parameters => parameters.Add("@Id", SqlDbType.BigInt).Value = id,
            command => One(command, ct), ct);

    private Task Save(string procedure, string actor, Action<SqlParameterCollection> fill, CancellationToken ct) =>
        Query(procedure, actor, parameters =>
        {
            fill(parameters);
            parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
        }, async command =>
        {
            await command.ExecuteNonQueryAsync(ct);
            return true;
        }, ct);

    private async Task<T> Query<T>(string procedure, string actor, Action<SqlParameterCollection> fill,
        Func<SqlCommand, Task<T>> read, CancellationToken ct)
    {
        try
        {
            await using var connection = connections.CreateConnection();
            await connection.OpenAsync(ct);
            await using var command = new SqlCommand(procedure, connection) { CommandType = CommandType.StoredProcedure };
            Text(command.Parameters, "@UsuarioId", actor, 128);
            fill(command.Parameters);
            return await read(command);
        }
        catch (SqlException error) when (error.Number is 2601 or 2627)
        {
            throw new MaintenanceException("El correo o nombre de perfil ya existe. Revisa los datos.");
        }
        catch (SqlException error) when (error.Message.Contains("[HYT_", StringComparison.Ordinal))
        {
            throw new MaintenanceException(error.Message[(error.Message.IndexOf(']') + 1)..].Trim());
        }
    }

    private static void Text(SqlParameterCollection parameters, string name, string? value, int size = -1) =>
        parameters.Add(name, SqlDbType.NVarChar, size).Value = (object?)value ?? DBNull.Value;

    private static async Task<JsonElement?> One(SqlCommand command, CancellationToken ct)
    {
        var raw = Convert.ToString(await command.ExecuteScalarAsync(ct));
        return string.IsNullOrWhiteSpace(raw) ? null : JsonSerializer.Deserialize<JsonElement>(raw);
    }

    private static async Task<SecurityPage> Page(SqlCommand command, int page, CancellationToken ct)
    {
        await using var reader = await command.ExecuteReaderAsync(ct);
        return await ReadPage(reader, page, ct);
    }

    private static async Task<SecurityPage> ReadPage(SqlDataReader reader, int page, CancellationToken ct)
    {
        var rows = await ReadRows(reader, ct);
        if (!await reader.NextResultAsync(ct) || !await reader.ReadAsync(ct))
            throw new InvalidOperationException("La consulta de seguridad no devolvió el total de registros.");
        return new SecurityPage(rows, reader.GetInt64(0), page);
    }

    private static async Task<IReadOnlyList<JsonElement>> ReadRows(SqlDataReader reader, CancellationToken ct)
    {
        if (!await reader.ReadAsync(ct) || reader.IsDBNull(0)) return [];
        return JsonSerializer.Deserialize<JsonElement[]>(reader.GetString(0)) ?? [];
    }
}
