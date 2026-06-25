using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Parametros;

public sealed class ParametroEmpresaRepository(IDbConnectionFactory connectionFactory) : IParametroEmpresaRepository
{
    public async Task<PagedResult<ParametroEmpresaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? tipoParametro, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<ParametroEmpresaDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarParametrosPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@TipoParametro", string.IsNullOrWhiteSpace(tipoParametro) ? (object)DBNull.Value : tipoParametro.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(MapParametro(reader));
        }

        return new PagedResult<ParametroEmpresaDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<ParametroEmpresaDto?> ObtenerAsync(int idEmpresa, int idParametroEmpresa, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ObtenerParametroEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdParametroEmpresa", idParametroEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        return await reader.ReadAsync(cancellationToken)
            ? MapParametro(reader)
            : null;
    }

    public async Task<ParametroEmpresaDto> GuardarAsync(GuardarParametroEmpresaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_GuardarParametroEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdParametroEmpresa", (object?)request.IdParametroEmpresa ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@TipoParametro", request.TipoParametro);
        command.Parameters.AddWithValue("@CodigoParametro", request.CodigoParametro);
        command.Parameters.AddWithValue("@ValorParametro", request.ValorParametro);
        command.Parameters.AddWithValue("@DescripcionParametro", request.DescripcionParametro);
        command.Parameters.AddWithValue("@FecIni", (object?)request.FecIni ?? DBNull.Value);
        command.Parameters.AddWithValue("@FecFin", (object?)request.FecFin ?? DBNull.Value);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el parametro grabado.");
        }

        return MapParametro(reader);
    }

    public async Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_CargarParametrosDefaultEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static ParametroEmpresaDto MapParametro(SqlDataReader reader)
    {
        return new ParametroEmpresaDto
        {
            IdParametroEmpresa = reader.GetInt32(reader.GetOrdinal("IdParametroEmpresa")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            TipoParametro = reader.GetString(reader.GetOrdinal("TipoParametro")),
            CodigoParametro = reader.GetString(reader.GetOrdinal("CodigoParametro")),
            ValorParametro = reader.GetString(reader.GetOrdinal("ValorParametro")),
            DescripcionParametro = reader.GetString(reader.GetOrdinal("DescripcionParametro")),
            FecIni = reader.IsDBNull(reader.GetOrdinal("FecIni")) ? null : reader.GetDateTime(reader.GetOrdinal("FecIni")),
            FecFin = reader.IsDBNull(reader.GetOrdinal("FecFin")) ? null : reader.GetDateTime(reader.GetOrdinal("FecFin")),
            Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
        };
    }
}
