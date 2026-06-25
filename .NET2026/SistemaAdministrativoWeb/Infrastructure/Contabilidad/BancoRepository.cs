using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BancoRepository(IDbConnectionFactory connectionFactory) : IBancoRepository
{
    public async Task<PagedResult<BancoDto>> ListarPaginadoAsync(string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = true, CancellationToken cancellationToken = default)
    {
        var result = new List<BancoDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarBancos", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@SoloActivos", soloActivos);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new BancoDto
            {
                IdBanco = reader.GetInt32(reader.GetOrdinal("IdBanco")),
                CodigoBanco = reader.GetString(reader.GetOrdinal("Codigo")),
                NombreBanco = reader.GetString(reader.GetOrdinal("Nombre")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return new PagedResult<BancoDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }
}
