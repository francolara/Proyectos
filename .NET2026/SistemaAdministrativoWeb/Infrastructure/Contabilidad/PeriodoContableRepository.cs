using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PeriodoContableRepository(IDbConnectionFactory connectionFactory) : IPeriodoContableRepository
{
    public async Task<PeriodoContableEstadoDto?> ObtenerAsync(int idEmpresa, string periodo, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerPeriodoContableEstado", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", (periodo ?? string.Empty).Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return Map(reader);
    }

    public async Task<PeriodoContableEstadoDto> GuardarAsync(GuardarPeriodoContableEstadoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarPeriodoContableEstado", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo.Trim());
        command.Parameters.AddWithValue("@Cerrado", request.Cerrado);
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el estado del periodo grabado.");
        }

        return Map(reader);
    }

    private static PeriodoContableEstadoDto Map(SqlDataReader reader)
    {
        return new PeriodoContableEstadoDto
        {
            IdPeriodoContableEstado = reader.IsDBNull(reader.GetOrdinal("IdPeriodoContableEstado")) ? null : reader.GetInt32(reader.GetOrdinal("IdPeriodoContableEstado")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
            Cerrado = reader.GetBoolean(reader.GetOrdinal("Cerrado")),
            FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
            UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro")),
            FechaCierre = reader.IsDBNull(reader.GetOrdinal("FechaCierre")) ? null : reader.GetDateTime(reader.GetOrdinal("FechaCierre")),
            UsuarioCierre = reader.IsDBNull(reader.GetOrdinal("UsuarioCierre")) ? null : reader.GetString(reader.GetOrdinal("UsuarioCierre")),
            FechaApertura = reader.IsDBNull(reader.GetOrdinal("FechaApertura")) ? null : reader.GetDateTime(reader.GetOrdinal("FechaApertura")),
            UsuarioApertura = reader.IsDBNull(reader.GetOrdinal("UsuarioApertura")) ? null : reader.GetString(reader.GetOrdinal("UsuarioApertura"))
        };
    }
}
