using System.Data;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;
namespace GestionServicios.Infrastructure.Followup;
public sealed class SqlDocumentFileRepository(SqlConnectionFactory connections):IDocumentFileRepository
{
    public async Task<DocumentFile?> GetAsync(DocumentKind kind,string actor,long fileId,CancellationToken ct)
    {
        _=kind.Procedure("Obtener");
        try
        {
            await using var connection=connections.CreateConnection();await connection.OpenAsync(ct);
            await using var command=new SqlCommand("dbo.spu_Seguimiento_ObtenerArchivo",connection){CommandType=CommandType.StoredProcedure};
            command.Parameters.Add("@UsuarioId",SqlDbType.NVarChar,128).Value=actor;command.Parameters.Add("@Tipo",SqlDbType.NVarChar,20).Value=kind.Code;command.Parameters.Add("@ArchivoId",SqlDbType.BigInt).Value=fileId;
            await using var reader=await command.ExecuteReaderAsync(ct);if(!await reader.ReadAsync(ct))return null;
            return new(reader.GetInt64(0),new(reader.GetString(1),reader.GetString(2),reader.GetString(3),reader.GetInt64(4),reader.GetString(5)));
        }
        catch(SqlException error)when(error.Message.Contains("[HYT_",StringComparison.Ordinal))
        {throw new MaintenanceException(error.Message[(error.Message.IndexOf(']')+1)..].Trim());}
    }
}