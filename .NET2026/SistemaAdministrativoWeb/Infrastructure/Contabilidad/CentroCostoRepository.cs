using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CentroCostoRepository(IDbConnectionFactory connectionFactory) : ICentroCostoRepository
{
    public async Task<IReadOnlyCollection<CentroCostoDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = new List<CentroCostoDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarCentroCostoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloActivos", soloActivos);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(MapCentroCosto(reader));
        }

        return result;
    }

    public async Task<PagedResult<CentroCostoDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = new List<CentroCostoDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarCentroCostoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloActivos", soloActivos);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(MapCentroCosto(reader));
        }

        return new PagedResult<CentroCostoDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<CentroCostoDto> GuardarAsync(GuardarCentroCostoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarCentroCostoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCentroCosto", (object?)request.IdCentroCosto ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@CodigoCentroCosto", request.CodigoCentroCosto);
        command.Parameters.AddWithValue("@NombreCentroCosto", request.NombreCentroCosto);
        command.Parameters.AddWithValue("@Estado", request.Estado);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el centro de costo grabado.");
        }

        return MapCentroCosto(reader);
    }

    public async Task EliminarAsync(int idEmpresa, int idCentroCosto, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarCentroCostoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdCentroCosto", idCentroCosto);
        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static CentroCostoDto MapCentroCosto(SqlDataReader reader)
    {
        return new CentroCostoDto
        {
            IdCentroCosto = reader.GetInt32(reader.GetOrdinal("IdCentroCosto")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            CodigoCentroCosto = reader.GetString(reader.GetOrdinal("CodigoCentroCosto")),
            NombreCentroCosto = reader.GetString(reader.GetOrdinal("NombreCentroCosto")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }
}
