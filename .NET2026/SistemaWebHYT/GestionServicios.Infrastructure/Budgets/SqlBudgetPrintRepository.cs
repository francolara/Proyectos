using System.Data;
using System.Security.Cryptography;
using System.Text;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Budgets;

public sealed class SqlBudgetPrintRepository(SqlConnectionFactory connections):IBudgetPrintRepository
{
    private async Task<T> Execute<T>(string procedure,string actor,Action<SqlParameterCollection> parameters,Func<SqlCommand,Task<T>> execute,CancellationToken ct)
    {
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand(procedure,connection){CommandType=CommandType.StoredProcedure};command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;parameters(command.Parameters);
            return await execute(command);
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal))
        {throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
    public Task<long> SaveAsync(string actor,long revisionId,byte[] expectedVersion,string format,string templateVersion,string html,CancellationToken ct)
        =>Execute("dbo.spu_Presupuesto_RegistrarImpresion",actor,p=>
        {
            p.Add("@CorrelacionId",SqlDbType.UniqueIdentifier).Value=Guid.NewGuid();p.Add("@RevisionId",SqlDbType.BigInt).Value=revisionId;p.Add("@VersionOriginal",SqlDbType.Binary,8).Value=expectedVersion;
            p.Add("@Formato",SqlDbType.NVarChar,10).Value=format;p.Add("@VersionFormato",SqlDbType.NVarChar,20).Value=templateVersion;p.Add("@ContenidoHTML",SqlDbType.NVarChar,-1).Value=html;
        },async command=>Convert.ToInt64(await command.ExecuteScalarAsync(ct)),ct);
    public Task<BudgetPrint?> GetAsync(string actor,long printId,CancellationToken ct,bool download=false)
        =>Execute<BudgetPrint?>("dbo.spu_Presupuesto_ObtenerImpresion",actor,p=>{p.Add("@ImpresionId",SqlDbType.BigInt).Value=printId;p.Add("@Descargar",SqlDbType.Bit).Value=download;},async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);if(!await reader.ReadAsync(ct)||reader.IsDBNull(7))return null;
            var html=reader.GetString(7);var hash=reader.GetString(8);var actual=Convert.ToHexString(SHA256.HashData(Encoding.Unicode.GetBytes(html)));
            if(!actual.Equals(hash,StringComparison.OrdinalIgnoreCase))throw new MaintenanceException("El contenido conservado no coincide con su huella de integridad.");
            return new(reader.GetInt64(0),reader.GetInt64(1),reader.GetInt64(2),reader.GetString(3),reader.GetString(4),new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(5),DateTimeKind.Utc)),reader.GetString(6),html,hash);
        },ct);
    public Task<BudgetPrintPage> ListAsync(string actor,long revisionId,int page,CancellationToken ct)
        =>Execute<BudgetPrintPage>("dbo.spu_Presupuesto_ListarImpresiones",actor,p=>
        {p.Add("@RevisionId",SqlDbType.BigInt).Value=revisionId;p.Add("@Pagina",SqlDbType.Int).Value=page;},async command=>
        {
            await using var reader=await command.ExecuteReaderAsync(ct);var rows=new List<BudgetPrintEntry>();
            while(await reader.ReadAsync(ct))rows.Add(new(reader.GetInt64(0),reader.GetInt64(1),reader.GetInt64(2),reader.GetString(3),reader.GetString(4),new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(5),DateTimeKind.Utc)),reader.GetString(6)));
            await reader.NextResultAsync(ct);await reader.ReadAsync(ct);return new(rows,reader.GetInt64(0),page);
        },ct);
}
