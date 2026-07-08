using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoCambioRepository(IDbConnectionFactory connectionFactory) : ITipoCambioRepository
{
    public async Task<IReadOnlyCollection<TipoCambioDto>> ListarPorCuentaAdministradoraAsync(int idCuentaAdministradora, short anio, byte mes, CancellationToken cancellationToken = default)
    {
        var result = new List<TipoCambioDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarTipoCambioPorCuentaAdministradora", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@Anio", anio);
        command.Parameters.AddWithValue("@Mes", mes);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(MapTipoCambio(reader));
        }

        return result;
    }

    public async Task<TipoCambioDto?> ObtenerAsync(int idTipoCambio, int idCuentaAdministradora, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerTipoCambio", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdTipoCambio", idTipoCambio);
        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return MapTipoCambio(reader);
    }

    public async Task<TipoCambioDto?> ObtenerPorFechaMonedaAsync(int idCuentaAdministradora, DateOnly fecha, string idMoneda, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerTipoCambioPorFecha", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaAdministradora", idCuentaAdministradora);
        command.Parameters.AddWithValue("@Fecha", fecha.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@IdMoneda", (idMoneda ?? string.Empty).Trim().ToUpperInvariant());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return null;
        }

        return MapTipoCambio(reader);
    }

    public async Task<TipoCambioDto> GuardarAsync(GuardarTipoCambioRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarTipoCambio", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdTipoCambio", (object?)request.IdTipoCambio ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdCuentaAdministradora", request.IdCuentaAdministradora);
        command.Parameters.AddWithValue("@Fecha", request.Fecha.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@Compra", request.Compra);
        command.Parameters.AddWithValue("@Venta", request.Venta);
        command.Parameters.AddWithValue("@CompraSBS", request.CompraSbs);
        command.Parameters.AddWithValue("@VentaSBS", request.VentaSbs);
        command.Parameters.AddWithValue("@Fuente", request.Fuente.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());
        command.Parameters.AddWithValue("@Estado", request.Estado);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el tipo de cambio grabado.");
        }

        return MapTipoCambio(reader);
    }

    private static TipoCambioDto MapTipoCambio(SqlDataReader reader)
    {
        return new TipoCambioDto
        {
            IdTipoCambio = reader.GetInt32(reader.GetOrdinal("IdTipoCambio")),
            IdCuentaAdministradora = reader.GetInt32(reader.GetOrdinal("IdCuentaAdministradora")),
            Fecha = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("Fecha"))),
            IdMoneda = NormalizarCodigoMoneda(reader.GetString(reader.GetOrdinal("IdMoneda"))),
            Compra = reader.GetDecimal(reader.GetOrdinal("Compra")),
            Venta = reader.GetDecimal(reader.GetOrdinal("Venta")),
            CompraSbs = reader.GetDecimal(reader.GetOrdinal("CompraSBS")),
            VentaSbs = reader.GetDecimal(reader.GetOrdinal("VentaSBS")),
            Fuente = reader.GetString(reader.GetOrdinal("Fuente")),
            UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro")),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        };
    }

    private static string NormalizarCodigoMoneda(string? idMoneda)
    {
        var valor = (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(valor))
        {
            return string.Empty;
        }

        var separadores = new[] { " - ", "-", " " };
        foreach (var separador in separadores)
        {
            var indice = valor.IndexOf(separador, StringComparison.Ordinal);
            if (indice > 0)
            {
                valor = valor[..indice].Trim();
                break;
            }
        }

        return valor.Length > 3 ? valor[..3] : valor;
    }
}
