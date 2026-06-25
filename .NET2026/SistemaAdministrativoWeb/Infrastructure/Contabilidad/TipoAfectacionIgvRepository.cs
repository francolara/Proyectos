using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoAfectacionIgvRepository(IDbConnectionFactory connectionFactory) : ITipoAfectacionIgvRepository
{
    public async Task<IReadOnlyCollection<TipoAfectacionIgvDto>> ListarActivosAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<TipoAfectacionIgvDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarTiposAfectacionIGV", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new TipoAfectacionIgvDto
            {
                IdTipoAfectacionIGV = reader.GetInt32(reader.GetOrdinal("IdTipoAfectacionIGV")),
                CodigoSunat = reader.GetString(reader.GetOrdinal("CodigoSunat")),
                NombreAfectacion = reader.GetString(reader.GetOrdinal("NombreAfectacion")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
