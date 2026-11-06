using System.Data;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Followup;

public sealed class SqlDocumentRepository(SqlConnectionFactory connections):IDocumentRepository
{
    private async Task<T> Execute<T>(DocumentKind kind,string operation,string actor,Action<SqlParameterCollection> parameters,Func<SqlCommand,Task<T>> action,CancellationToken ct)
    {
        var procedure=kind.Procedure(operation);
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand(procedure,connection){CommandType=CommandType.StoredProcedure};
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;parameters(command.Parameters);
            return await action(command);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal))
        {throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
        catch(SqlException error)when(error.Number is 8115 or 2628 or 8152)
        {throw new MaintenanceException("Un importe o texto excede su longitud permitida. Revisa los datos.");}
    }
    private static string? Text(SqlDataReader reader,int index)=>reader.IsDBNull(index)?null:reader.GetString(index);
    public Task<DocumentPage> ListAsync(DocumentKind kind,string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,bool? confirmed,CancellationToken ct,string? client=null,string? store=null)
        =>Execute(kind,"Listar",actor,p=>
        {
            p.Add("@Desde",SqlDbType.DateTime2).Value=from.UtcDateTime;p.Add("@HastaExclusiva",SqlDbType.DateTime2).Value=until.UtcDateTime;
            p.Add("@Cliente",SqlDbType.NVarChar,200).Value=(object?)client??DBNull.Value;p.Add("@Tienda",SqlDbType.NVarChar,200).Value=(object?)store??DBNull.Value;
            p.Add("@Busqueda",SqlDbType.NVarChar,200).Value=(object?)search??DBNull.Value;p.Add("@Pagina",SqlDbType.Int).Value=page;
            p.Add("@Estado",SqlDbType.NVarChar,40).Value=(object?)state??DBNull.Value;p.Add("@Confirmada",SqlDbType.Bit).Value=(object?)confirmed??DBNull.Value;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);await reader.ReadAsync(ct);
            var rows=JsonSerializer.Deserialize<JsonElement[]>(reader.IsDBNull(0)?"[]":reader.GetString(0))??[];
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new DocumentPage(rows,reader.GetInt64(0),page);
        },ct);
    public Task<DocumentRecord?> GetAsync(DocumentKind kind,string actor,long? id,long? parentId,int auditPage,CancellationToken ct)
        =>Execute<DocumentRecord?>(kind,"Obtener",actor,p=>
        {
            p.Add("@Id",SqlDbType.BigInt).Value=(object?)id??DBNull.Value;p.Add("@PadreId",SqlDbType.BigInt).Value=(object?)parentId??DBNull.Value;
            p.Add("@PaginaAuditoria",SqlDbType.Int).Value=auditPage;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);
            if(!await reader.ReadAsync(ct)||reader.IsDBNull(0))return null;
            var data=JsonSerializer.Deserialize<JsonElement>(reader.GetString(0));var audit=new List<BudgetAudit>();
            await reader.NextResultAsync(ct);
            while(await reader.ReadAsync(ct))audit.Add(new(reader.GetInt64(0),reader.GetString(1),Text(reader,2),Text(reader,3),new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(4),DateTimeKind.Utc)),reader.IsDBNull(5)?null:DateOnly.FromDateTime(reader.GetDateTime(5)),reader.GetString(6),Text(reader,7),Text(reader,8)));
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new DocumentRecord(data,audit,reader.GetInt64(0));
        },ct);
    public Task<DocumentSaved> SaveAsync(DocumentKind kind,string actor,long parentId,byte[] parentVersion,string payload,Guid operationId,long? id,byte[]? originalVersion,CancellationToken ct)
        =>Execute(kind,"Guardar",actor,p=>
        {
            p.Add("@PadreId",SqlDbType.BigInt).Value=parentId;p.Add("@VersionPadre",SqlDbType.Binary,8).Value=parentVersion;
            p.Add("@Datos",SqlDbType.NVarChar,-1).Value=payload;p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=operationId;
            p.Add("@Id",SqlDbType.BigInt).Value=(object?)id??DBNull.Value;p.Add("@VersionOriginal",SqlDbType.Binary,8).Value=(object?)originalVersion??DBNull.Value;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);await reader.ReadAsync(ct);
            return new DocumentSaved(reader.GetInt64(0),reader.IsDBNull(1)?null:reader.GetInt64(1));
        },ct);
    public Task ChangeStateAsync(DocumentKind kind,string actor,long id,long parentId,byte[] parentVersion,byte[] originalVersion,Guid operation,string state,DateOnly eventDate,string responsible,string? reason,CancellationToken ct)
        =>Execute(kind,"CambiarEstado",actor,p=>
        {
            p.Add("@Id",SqlDbType.BigInt).Value=id;p.Add("@PadreId",SqlDbType.BigInt).Value=parentId;
            p.Add("@VersionPadre",SqlDbType.Binary,8).Value=parentVersion;p.Add("@VersionOriginal",SqlDbType.Binary,8).Value=originalVersion;
            p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=operation;p.Add("@NuevoEstado",SqlDbType.NVarChar,40).Value=state;
            p.Add("@FechaEvento",SqlDbType.Date).Value=eventDate.ToDateTime(TimeOnly.MinValue);p.Add("@Responsable",SqlDbType.NVarChar,150).Value=responsible;p.Add("@Motivo",SqlDbType.NVarChar,2000).Value=(object?)reason??DBNull.Value;
        },async command=>Convert.ToInt64(await command.ExecuteScalarAsync(ct)),ct);}