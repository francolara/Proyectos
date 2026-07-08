using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoPercepcionRepository(IDbConnectionFactory connectionFactory) : ITipoPercepcionRepository
{
    public async Task<IReadOnlyCollection<TipoPercepcionDto>> ListarActivasAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<TipoPercepcionDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarTiposPercepcion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new TipoPercepcionDto
            {
                IdTipoPercepcion = reader.GetInt32(reader.GetOrdinal("IdTipoPercepcion")),
                Codigo = reader.GetString(reader.GetOrdinal("Codigo")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                Porcentaje = reader.GetDecimal(reader.GetOrdinal("Porcentaje")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
