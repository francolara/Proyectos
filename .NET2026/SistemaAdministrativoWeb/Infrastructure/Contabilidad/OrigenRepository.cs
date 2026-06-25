using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class OrigenRepository(IDbConnectionFactory connectionFactory) : IOrigenRepository
{
    public async Task<IReadOnlyCollection<OrigenDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = new List<OrigenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarOrigenesActivos", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@SoloActivos", soloActivos);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new OrigenDto
            {
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                ModuloOrigen = reader.GetString(reader.GetOrdinal("ModuloOrigen")),
                PermiteRegistroManual = reader.GetBoolean(reader.GetOrdinal("PermiteRegistroManual")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }

    public async Task<PagedResult<OrigenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default)
    {
        var result = new List<OrigenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarOrigenesActivos", connection)
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
            result.Add(new OrigenDto
            {
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                ModuloOrigen = reader.GetString(reader.GetOrdinal("ModuloOrigen")),
                PermiteRegistroManual = reader.GetBoolean(reader.GetOrdinal("PermiteRegistroManual")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return new PagedResult<OrigenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<OrigenDto> GuardarAsync(GuardarOrigenRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarOrigenPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdOrigen", (object?)request.IdOrigen ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@CodigoOrigen", request.CodigoOrigen);
        command.Parameters.AddWithValue("@NombreOrigen", request.NombreOrigen);
        command.Parameters.AddWithValue("@ModuloOrigen", request.ModuloOrigen);
        command.Parameters.AddWithValue("@PermiteRegistroManual", request.PermiteRegistroManual);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el origen grabado.");
        }

        return new OrigenDto
        {
            IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
            CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
            NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
            ModuloOrigen = reader.GetString(reader.GetOrdinal("ModuloOrigen")),
            PermiteRegistroManual = reader.GetBoolean(reader.GetOrdinal("PermiteRegistroManual")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }

    public async Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_CargarOrigenesDefaultEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}
