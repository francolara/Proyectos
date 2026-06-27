using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class DetraccionSunatRepository(IDbConnectionFactory connectionFactory) : IDetraccionSunatRepository
{
    public async Task<IReadOnlyCollection<DetraccionSunatDto>> ListarActivasAsync(CancellationToken cancellationToken = default)
    {
        var result = new List<DetraccionSunatDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarDetraccionesSunat", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new DetraccionSunatDto
            {
                IdDetraccionSunat = reader.GetInt32(reader.GetOrdinal("IdDetraccionSunat")),
                CodigoSunat = reader.GetString(reader.GetOrdinal("CodigoSunat")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                Porcentaje = reader.GetDecimal(reader.GetOrdinal("Porcentaje")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
