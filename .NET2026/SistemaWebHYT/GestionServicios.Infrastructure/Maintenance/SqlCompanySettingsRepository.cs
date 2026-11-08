using System.Data;
using System.Text.Json;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Maintenance;

public sealed class SqlCompanySettingsRepository(SqlConnectionFactory connections):ICompanySettingsRepository
{
    public async Task<RecordSnapshot?> GetAsync(string actor,int page,CancellationToken ct)
    {
        await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
        await using var command=new SqlCommand("dbo.spu_ConfiguracionHYT_Obtener",connection){CommandType=CommandType.StoredProcedure};
        command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;command.Parameters.Add("@Pagina",SqlDbType.Int).Value=page;
        await using var reader=await command.ExecuteReaderAsync(ct);
        if(!await reader.ReadAsync(ct)||reader.IsDBNull(0)||string.IsNullOrWhiteSpace(reader.GetString(0)))return null;
        var values=JsonSerializer.Deserialize<Dictionary<string,JsonElement>>(reader.GetString(0))!;var audit=new List<AuditEntry>();
        await reader.NextResultAsync(ct);while(await reader.ReadAsync(ct))audit.Add(new(reader.GetString(0),reader.GetString(1),new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(2),DateTimeKind.Utc)),reader.IsDBNull(3)?null:reader.GetString(3),reader.IsDBNull(4)?null:reader.GetString(4),reader.IsDBNull(5)?null:reader.GetString(5)));
        await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new(1,values,[],audit,reader.GetInt64(0));
    }
    public async Task SaveSignatureAsync(string actor,string payload,byte[] expectedVersion,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand("dbo.spu_ConfiguracionHYT_FirmaGuardar",connection){CommandType=CommandType.StoredProcedure};
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;
            command.Parameters.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();
            command.Parameters.Add("@Datos",SqlDbType.NVarChar,2000).Value=payload;
            command.Parameters.Add("@VersionOriginal",SqlDbType.Binary,8).Value=expectedVersion;
            await command.ExecuteNonQueryAsync(ct);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal))
        {throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
    public async Task SaveAsync(string actor,string payload,byte[]? expectedVersion,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand("dbo.spu_ConfiguracionHYT_Guardar",connection){CommandType=CommandType.StoredProcedure};
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;command.Parameters.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();
            command.Parameters.Add("@Datos",SqlDbType.NVarChar,-1).Value=payload;command.Parameters.Add("@VersionOriginal",SqlDbType.Binary,8).Value=(object?)expectedVersion??DBNull.Value;
            await command.ExecuteNonQueryAsync(ct);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal))
        {throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
}
