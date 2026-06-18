using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoComprobanteRepository(IDbConnectionFactory connectionFactory) : ITipoComprobanteRepository
{
    public async Task<IReadOnlyCollection<TipoComprobanteDto>> ListarActivosAsync(bool usoCompras, bool usoVentas, CancellationToken cancellationToken = default)
    {
        var result = new List<TipoComprobanteDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarTiposComprobanteActivos", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@UsoCompras", usoCompras);
        command.Parameters.AddWithValue("@UsoVentas", usoVentas);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new TipoComprobanteDto
            {
                IdTipoComprobante = reader.GetInt32(reader.GetOrdinal("IdTipoComprobante")),
                CodigoTipoComprobante = reader.GetString(reader.GetOrdinal("CodigoTipoComprobante")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                UsoCompras = reader.GetBoolean(reader.GetOrdinal("UsoCompras")),
                UsoVentas = reader.GetBoolean(reader.GetOrdinal("UsoVentas")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
