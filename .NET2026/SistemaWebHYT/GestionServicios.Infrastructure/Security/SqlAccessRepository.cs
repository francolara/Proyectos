using System.Data;
using GestionServicios.Application.Security;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Security;

public sealed class SqlAccessRepository(SqlConnectionFactory connections) : IAccessRepository
{
    public Task<AccessAccount?> FindByNameAsync(string normalizedName, CancellationToken cancellationToken = default)
        => ReadAsync(null, normalizedName, cancellationToken);

    public Task<AccessAccount?> FindByIdAsync(string id, CancellationToken cancellationToken = default)
        => ReadAsync(id, null, cancellationToken);

    private async Task<AccessAccount?> ReadAsync(string? id, string? name, CancellationToken cancellationToken)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand("dbo.spu_ConsultarAcceso", connection) { CommandType = CommandType.StoredProcedure };
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = (object?)id ?? DBNull.Value;
        command.Parameters.Add("@NombreNormalizado", SqlDbType.NVarChar, 256).Value = (object?)name ?? DBNull.Value;
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) return null;
        var userId = reader.GetString(0);
        var userName = reader.GetString(1);
        var displayName = reader.GetString(2);
        var hash = reader.IsDBNull(3) ? null : reader.GetString(3);
        var stamp = reader.GetString(4);
        var active = reader.GetBoolean(5);
        var lockout = reader.IsDBNull(6) ? (DateTimeOffset?)null : reader.GetFieldValue<DateTimeOffset>(6);
        var roles = new List<string>();
        var permissions = new List<string>();
        await reader.NextResultAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) roles.Add(reader.GetString(0));
        await reader.NextResultAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) permissions.Add(reader.GetString(0));
        return new(userId, userName, displayName, hash, stamp, active, lockout, roles, permissions);
    }

    public async Task<bool> CompleteAttemptAsync(string id, string expectedStamp, bool passwordValid,
        string? upgradedHash, Guid correlationId, CancellationToken cancellationToken = default)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand("dbo.spu_FinalizarAcceso", connection) { CommandType = CommandType.StoredProcedure };
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = id;
        command.Parameters.Add("@SelloEsperado", SqlDbType.NVarChar, -1).Value = expectedStamp;
        command.Parameters.Add("@PasswordValido", SqlDbType.Bit).Value = passwordValid;
        command.Parameters.Add("@HashActualizado", SqlDbType.NVarChar, -1).Value = (object?)upgradedHash ?? DBNull.Value;
        command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = correlationId;
        return Convert.ToBoolean(await command.ExecuteScalarAsync(cancellationToken));
    }

    public async Task RecordLogoutAsync(string id, Guid correlationId, CancellationToken cancellationToken = default)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(cancellationToken);
        await using var command = new SqlCommand("dbo.spu_RegistrarSalida", connection) { CommandType = CommandType.StoredProcedure };
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = id;
        command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = correlationId;
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}
