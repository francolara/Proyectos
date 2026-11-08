using System.Data;
using System.Text.Json;
using GestionServicios.Application.Reporting;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Reporting;
public sealed class SqlReportingRepository(SqlConnectionFactory connections):IReportingRepository
{
    private async Task<T> Query<T>(string procedure,string actor,DateTimeOffset from,DateTimeOffset until,Action<SqlParameterCollection> parameters,Func<SqlCommand,Task<T>> action,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);await using var command=new SqlCommand(procedure,connection){CommandType=CommandType.StoredProcedure};
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;command.Parameters.Add("@Desde",SqlDbType.DateTime2).Value=from.UtcDateTime;command.Parameters.Add("@HastaExclusiva",SqlDbType.DateTime2).Value=until.UtcDateTime;parameters(command.Parameters);return await action(command);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal)){throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
    public Task<JsonElement> DashboardAsync(string actor,DateTimeOffset from,DateTimeOffset until,CancellationToken ct,string? client=null,string? store=null,int queuePage=1)
        =>Query("dbo.spu_Dashboard_Obtener",actor,from,until,p=>{p.Add("@Cliente",SqlDbType.NVarChar,200).Value=(object?)client??DBNull.Value;p.Add("@Tienda",SqlDbType.NVarChar,200).Value=(object?)store??DBNull.Value;p.Add("@PaginaPendientes",SqlDbType.Int).Value=queuePage;},async command=>JsonSerializer.Deserialize<JsonElement>(Convert.ToString(await command.ExecuteScalarAsync(ct))!),ct);
    public Task<ReportPage> ReportAsync(string actor,string kind,DateTimeOffset from,DateTimeOffset until,int page,CancellationToken ct,string? client=null,string? store=null,string? situation=null,ReportFilter? filters=null)
    {
        if(!ReportKinds.All.ContainsKey(kind)&&kind is not ("GANADOS" or "OC_COMPLETADAS" or "OC_CON_ACTA" or "OC_PENDIENTES_FACTURA"))throw new MaintenanceException("Reporte no válido.");
        return Query("dbo.spu_Reporte_Listar",actor,from,until,p=>{p.Add("@Cliente",SqlDbType.NVarChar,200).Value=(object?)client??DBNull.Value;p.Add("@Tienda",SqlDbType.NVarChar,200).Value=(object?)store??DBNull.Value;p.Add("@Situacion",SqlDbType.NVarChar,30).Value=(object?)situation??DBNull.Value;p.Add("@Estado",SqlDbType.NVarChar,40).Value=(object?)filters?.State??DBNull.Value;p.Add("@TipoAtencion",SqlDbType.NVarChar,20).Value=(object?)filters?.Type??DBNull.Value;p.Add("@Busqueda",SqlDbType.NVarChar,200).Value=(object?)filters?.Search??DBNull.Value;
            p.Add("@PendienteGuia",SqlDbType.Bit).Value=filters?.PendingGuide??false;p.Add("@PendienteOC",SqlDbType.Bit).Value=filters?.PendingOrder??false;p.Add("@PendienteConfirmacion",SqlDbType.Bit).Value=filters?.PendingConfirmation??false;p.Add("@PendienteFacturar",SqlDbType.Bit).Value=filters?.PendingInvoice??false;p.Add("@Exportar",SqlDbType.Bit).Value=filters?.Export??false;p.Add("@Reporte",SqlDbType.NVarChar,40).Value=kind;p.Add("@Pagina",SqlDbType.Int).Value=page;},async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);await reader.ReadAsync(ct);var rows=JsonSerializer.Deserialize<JsonElement[]>(reader.IsDBNull(0)?"[]":reader.GetString(0))??[];
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);var total=reader.GetInt64(0);await reader.NextResultAsync(ct);var amounts=new List<ReportAmount>();
            while(await reader.ReadAsync(ct))amounts.Add(new(reader.GetString(0),reader.IsDBNull(1)?null:reader.GetDecimal(1),reader.GetInt64(2)));
            return new ReportPage(rows,total,amounts,page);
        },ct);
    }
}