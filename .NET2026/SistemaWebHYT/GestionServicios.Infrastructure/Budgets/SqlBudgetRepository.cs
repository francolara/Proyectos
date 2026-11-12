using System.Data;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Budgets;

public sealed class SqlBudgetRepository(SqlConnectionFactory connections) : IBudgetRepository
{
    public Task<JsonElement> DefaultsAsync(string actor,DateOnly date,long? clientId,long? storeId,CancellationToken ct)
        =>Execute("dbo.spu_Presupuesto_DatosIniciales",actor,p=>
        {
            p.Add("@Fecha",SqlDbType.Date).Value=date.ToDateTime(TimeOnly.MinValue);p.Add("@ClienteId",SqlDbType.BigInt).Value=(object?)clientId??DBNull.Value;
            p.Add("@TiendaId",SqlDbType.BigInt).Value=(object?)storeId??DBNull.Value;
        },async command=>
        {
            var json=Convert.ToString(await command.ExecuteScalarAsync(ct));return JsonSerializer.Deserialize<JsonElement>(json??"{}");
        },ct);
    private async Task<T> Execute<T>(string procedure,string actor,Action<SqlParameterCollection> parameters,Func<SqlCommand,Task<T>> execute,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection(); await connection.OpenAsync(ct);
            await using var command=new SqlCommand(procedure,connection) { CommandType=CommandType.StoredProcedure };
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;
            parameters(command.Parameters); return await execute(command);
        }
        catch (SqlException error) when (error.Message.Contains("[HYT_",StringComparison.Ordinal))
        { throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim()); }
        catch (SqlException error) when (error.Number is 8115 or 2628 or 8152)
        { throw new MaintenanceException("Un importe o texto excede su longitud permitida. Revisa los datos."); }
    }
    private static string? Text(SqlDataReader reader,int index) => reader.IsDBNull(index)?null:reader.GetString(index);
    private static decimal? Money(SqlDataReader reader,int index) => reader.IsDBNull(index)?null:reader.GetDecimal(index);
    private static long? Id(SqlDataReader reader,int index) => reader.IsDBNull(index)?null:Convert.ToInt64(reader.GetValue(index));
    private static DateTimeOffset Timestamp(SqlDataReader reader,int index) => new(DateTime.SpecifyKind(reader.GetDateTime(index),DateTimeKind.Utc));
    private static void Original(SqlParameterCollection p,BudgetVersion? version)
    {
        p.Add("@PresupuestoId",SqlDbType.BigInt).Value=(object?)version?.BudgetId??DBNull.Value;
        p.Add("@RevisionId",SqlDbType.BigInt).Value=(object?)version?.RevisionId??DBNull.Value;
        p.Add("@VersionOriginal",SqlDbType.Binary,8).Value=(object?)version?.RevisionVersion??DBNull.Value;
        p.Add("@VersionPresupuesto",SqlDbType.Binary,8).Value=(object?)version?.HeaderVersion??DBNull.Value;
    }
    public Task<BudgetPage> ListAsync(string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,CancellationToken ct,string? client=null,string? store=null,string? type=null,bool pendingGuide=false,bool pendingOrder=false)
        => Execute("dbo.spu_Presupuesto_Listar",actor,p=>
        {
            p.Add("@Desde",SqlDbType.DateTime2).Value=from.UtcDateTime;p.Add("@HastaExclusiva",SqlDbType.DateTime2).Value=until.UtcDateTime;
            p.Add("@Cliente",SqlDbType.NVarChar,200).Value=(object?)client??DBNull.Value;p.Add("@Tienda",SqlDbType.NVarChar,200).Value=(object?)store??DBNull.Value;
            p.Add("@Busqueda",SqlDbType.NVarChar,200).Value=(object?)search??DBNull.Value;p.Add("@Pagina",SqlDbType.Int).Value=page;
            p.Add("@Estado",SqlDbType.NVarChar,20).Value=(object?)state??DBNull.Value;
            p.Add("@TipoAtencion",SqlDbType.NVarChar,20).Value=(object?)type??DBNull.Value;
            p.Add("@PendienteGuia",SqlDbType.Bit).Value=pendingGuide;
            p.Add("@PendienteOC",SqlDbType.Bit).Value=pendingOrder;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);var rows=new List<BudgetListItem>();
            while(await reader.ReadAsync(ct)) rows.Add(new(reader.GetInt64(0),reader.GetString(1),reader.GetString(2),reader.GetString(3),reader.GetInt64(4),reader.GetInt64(5),reader.GetInt64(6),reader.GetInt32(7),
                reader.GetString(8),reader.GetString(9),reader.GetString(10),reader.GetString(11),Money(reader,12),Money(reader,13),Money(reader,14),reader.IsDBNull(15)?null:DateOnly.FromDateTime(reader.GetDateTime(15)),
                Timestamp(reader,16),reader.GetString(17),reader.IsDBNull(18)?null:Timestamp(reader,18),Text(reader,19),Id(reader,20),Text(reader,21),Id(reader,22),Text(reader,23),Text(reader,24),reader.GetInt64(25),Id(reader,26),reader.FieldCount>27&&reader.GetBoolean(27),reader.FieldCount>28?Text(reader,28):null));
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);var total=reader.GetInt64(0);
            var counts=new Dictionary<string,long>();
            if(await reader.NextResultAsync(ct))while(await reader.ReadAsync(ct))counts[reader.GetString(0)]=reader.GetInt64(1);
            return new BudgetPage(rows,total,page,counts);
        },ct);
    public Task<BudgetRecord?> GetAsync(string actor,long id,long? revisionId,int auditPage,CancellationToken ct)
        => Execute<BudgetRecord?>("dbo.spu_Presupuesto_Obtener",actor,p=>
        {
            p.Add("@Id",SqlDbType.BigInt).Value=id;p.Add("@RevisionId",SqlDbType.BigInt).Value=(object?)revisionId??DBNull.Value;
            p.Add("@PaginaAuditoria",SqlDbType.Int).Value=auditPage;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);
            if(!await reader.ReadAsync(ct)||reader.IsDBNull(0)||string.IsNullOrWhiteSpace(reader.GetString(0)))return null;
            var data=JsonSerializer.Deserialize<JsonElement>(reader.GetString(0));
            var audit=new List<BudgetAudit>();await reader.NextResultAsync(ct);
            while(await reader.ReadAsync(ct)) audit.Add(new(reader.GetInt64(0),reader.GetString(1),Text(reader,2),Text(reader,3),Timestamp(reader,4),reader.IsDBNull(5)?null:DateOnly.FromDateTime(reader.GetDateTime(5)),reader.GetString(6),Text(reader,7),Text(reader,8),reader.FieldCount>9?Text(reader,9):null));
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new BudgetRecord(data,audit,reader.GetInt64(0));
        },ct);
    public Task<BudgetVersion> SaveAsync(string actor,string payload,BudgetVersion? original,bool newRevision,CancellationToken ct)
        => Execute("dbo.spu_Presupuesto_Guardar",actor,p=>
        {
            Original(p,original);p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();
            p.Add("@Datos",SqlDbType.NVarChar,-1).Value=payload;p.Add("@NuevaRevision",SqlDbType.Bit).Value=newRevision;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);await reader.ReadAsync(ct);
            return new BudgetVersion(reader.GetInt64(0),reader.GetInt64(1),reader.GetString(2),reader.GetInt32(3),(byte[])reader.GetValue(4),(byte[])reader.GetValue(5));
        },ct);
    public async Task ChangeStateAsync(string actor,BudgetVersion original,string newState,DateOnly eventDate,string responsible,string? reason,CancellationToken ct)
        => _=await Execute("dbo.spu_Presupuesto_CambiarEstado",actor,p=>
        {
            Original(p,original);p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();
            p.Add("@NuevoEstado",SqlDbType.NVarChar,20).Value=newState;p.Add("@FechaEvento",SqlDbType.Date).Value=eventDate.ToDateTime(TimeOnly.MinValue);
            p.Add("@Responsable",SqlDbType.NVarChar,150).Value=responsible;p.Add("@Motivo",SqlDbType.NVarChar,2000).Value=(object?)reason??DBNull.Value;
        },async command=>{await command.ExecuteNonQueryAsync(ct);return true;},ct);
    public Task<TariffChoicePage> TariffsAsync(string actor,long clientId,DateOnly date,string currency,string? search,int page,CancellationToken ct,long? materialId=null,long? tariffLineId=null)
        => Execute("dbo.spu_Presupuesto_AyudaTarifario",actor,p=>
        {
            p.Add("@ClienteId",SqlDbType.BigInt).Value=clientId;p.Add("@Fecha",SqlDbType.Date).Value=date.ToDateTime(TimeOnly.MinValue);
            p.Add("@CodigoMoneda",SqlDbType.NVarChar,3).Value=currency;p.Add("@Busqueda",SqlDbType.NVarChar,200).Value=(object?)search??DBNull.Value;p.Add("@Pagina",SqlDbType.Int).Value=page;p.Add("@MaterialId",SqlDbType.BigInt).Value=(object?)materialId??DBNull.Value;p.Add("@PartidaId",SqlDbType.BigInt).Value=(object?)tariffLineId??DBNull.Value;
        },async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);var rows=new List<TariffChoice>();
            while(await reader.ReadAsync(ct)) rows.Add(new(reader.GetInt64(0),reader.GetInt64(1),reader.GetString(2),Text(reader,3),Text(reader,4),reader.GetString(5),reader.GetString(6),reader.GetString(7),reader.GetInt32(8),
                reader.GetString(9),reader.GetString(10),Money(reader,11),Text(reader,12),Text(reader,13),Text(reader,14)));
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new TariffChoicePage(rows,reader.GetInt64(0),page);
        },ct);
}
