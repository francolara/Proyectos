using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public sealed class EmpresaRepository(IDbConnectionFactory connectionFactory) : IEmpresaRepository
{
    public async Task<IReadOnlyCollection<EmpresaDisponibleDto>> ListarPorUsuarioAsync(string aspNetUserId, CancellationToken cancellationToken = default)
    {
        var empresas = new List<EmpresaDisponibleDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_SEG_ListarEmpresasPorUsuario", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@AspNetUserId", aspNetUserId);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            empresas.Add(new EmpresaDisponibleDto
            {
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                CodigoEmpresa = reader.GetString(reader.GetOrdinal("CodigoEmpresa")),
                RazonSocial = reader.GetString(reader.GetOrdinal("RazonSocial")),
                NombreComercial = reader.IsDBNull(reader.GetOrdinal("NombreComercial"))
                    ? null
                    : reader.GetString(reader.GetOrdinal("NombreComercial")),
                Ruc = reader.GetString(reader.GetOrdinal("Ruc")),
                EsEmpresaPredeterminada = reader.GetBoolean(reader.GetOrdinal("EsEmpresaPredeterminada"))
            });
        }

        return empresas;
    }
}
