using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MonedaRepository(IDbConnectionFactory connectionFactory) : IMonedaRepository
{
    public async Task<IReadOnlyCollection<MonedaDto>> ListarActivasAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<MonedaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarMonedasActivas", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new MonedaDto
            {
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                NombreMoneda = reader.GetString(reader.GetOrdinal("NombreMoneda")),
                SimboloMoneda = reader.GetString(reader.GetOrdinal("SimboloMoneda")),
                EsMonedaBase = reader.GetBoolean(reader.GetOrdinal("EsMonedaBase")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
