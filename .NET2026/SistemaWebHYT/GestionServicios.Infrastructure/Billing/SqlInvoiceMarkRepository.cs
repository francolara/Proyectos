using System.Data;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Billing;
public sealed class SqlInvoiceMarkRepository(SqlConnectionFactory connections):IInvoiceMarkRepository
{
    private async Task<T> Execute<T>(string procedure,string actor,Action<SqlParameterCollection> parameters,Func<SqlCommand,Task<T>> action,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand(procedure,connection){CommandType=CommandType.StoredProcedure};command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;parameters(command.Parameters);return await action(command);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal)){throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
    public Task<JsonElement?> GetAsync(string actor,long orderId,int page,CancellationToken ct)
        =>Execute<JsonElement?>("dbo.spu_Facturado_Obtener",actor,p=>{p.Add("@OrdenCompraId",SqlDbType.BigInt).Value=orderId;p.Add("@Pagina",SqlDbType.Int).Value=page;},async command=>
        {var text=Convert.ToString(await command.ExecuteScalarAsync(ct));return string.IsNullOrWhiteSpace(text)?null:JsonSerializer.Deserialize<JsonElement>(text);},ct);
    public Task<long> SaveAsync(string actor,InvoiceMarkCommand input,CancellationToken ct)
    {
        if(input.OrderId<=0||input.OrderVersion.Length!=8||input.OperationId==Guid.Empty||input.ReferenceId is <=0||input.ReferenceId.HasValue&&input.ReferenceVersion?.Length!=8)
            throw new MaintenanceException("Vuelve a abrir la referencia del documento.");
        if(input.Withdraw&&!input.ReferenceId.HasValue||!input.Withdraw&&(input.Date is null||input.Type is not("01" or "03")||string.IsNullOrWhiteSpace(input.Series)||string.IsNullOrWhiteSpace(input.Number))||input.Series?.Length>30||input.Number?.Length>50||input.Reason?.Length>1000||input.ReferenceId.HasValue&&string.IsNullOrWhiteSpace(input.Reason))
            throw new MaintenanceException("Revisa fecha, tipo, serie, número y motivo de la referencia.");
        return Execute("dbo.spu_Facturado_Guardar",actor,p=>
        {
            p.Add("@OrdenCompraId",SqlDbType.BigInt).Value=input.OrderId;p.Add("@VersionOC",SqlDbType.Binary,8).Value=input.OrderVersion;p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=input.OperationId;
            p.Add("@FechaDocumento",SqlDbType.Date).Value=input.Date.HasValue?input.Date.Value.ToDateTime(TimeOnly.MinValue):DBNull.Value;
            p.Add("@CodigoTipoDocumento",SqlDbType.NVarChar,2).Value=(object?)input.Type??DBNull.Value;p.Add("@Serie",SqlDbType.NVarChar,30).Value=(object?)input.Series?.Trim()??DBNull.Value;p.Add("@NumeroDocumento",SqlDbType.NVarChar,50).Value=(object?)input.Number?.Trim()??DBNull.Value;
            p.Add("@ReferenciaId",SqlDbType.BigInt).Value=(object?)input.ReferenceId??DBNull.Value;p.Add("@VersionReferencia",SqlDbType.Binary,8).Value=(object?)input.ReferenceVersion??DBNull.Value;
            p.Add("@Retirar",SqlDbType.Bit).Value=input.Withdraw;p.Add("@Motivo",SqlDbType.NVarChar,1000).Value=(object?)input.Reason?.Trim()??DBNull.Value;
        },async command=>Convert.ToInt64(await command.ExecuteScalarAsync(ct)),ct);
    }
}