using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ClienteRepository(IDbConnectionFactory connectionFactory) : IClienteRepository
{
    public async Task<IReadOnlyCollection<ClienteDto>> ListarActivosPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        var result = new List<ClienteDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarClientesActivosPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new ClienteDto
            {
                IdCliente = reader.GetInt32(reader.GetOrdinal("IdCliente")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
                CodigoCliente = reader.GetString(reader.GetOrdinal("CodigoCliente")),
                TipoDocumento = reader.GetString(reader.GetOrdinal("TipoDocumento")),
                NumeroDocumento = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreCompleto = reader.GetString(reader.GetOrdinal("NombreCompleto")),
                CorreoElectronico = reader.IsDBNull(reader.GetOrdinal("CorreoElectronico")) ? null : reader.GetString(reader.GetOrdinal("CorreoElectronico")),
                Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                LimiteCredito = reader.GetDecimal(reader.GetOrdinal("LimiteCredito")),
                DiasCredito = reader.GetInt32(reader.GetOrdinal("DiasCredito")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
