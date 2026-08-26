using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiarioRepository(IDbConnectionFactory connectionFactory) : ILibroDiarioRepository
{
    public async Task<IReadOnlyCollection<LibroDiarioItemDto>> ListarAsync(LibroDiarioRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<LibroDiarioItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ReporteLibroDiario", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo);
        command.Parameters.AddWithValue("@Moneda", "PEN");
        command.Parameters.AddWithValue("@Modo", string.IsNullOrWhiteSpace(request.Modo) ? "A" : request.Modo.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@CuentaDesde", string.IsNullOrWhiteSpace(request.CuentaDesde) ? (object)DBNull.Value : request.CuentaDesde.Trim());
        command.Parameters.AddWithValue("@CuentaHasta", string.IsNullOrWhiteSpace(request.CuentaHasta) ? (object)DBNull.Value : request.CuentaHasta.Trim());
        command.Parameters.AddWithValue("@OrigenDesde", string.IsNullOrWhiteSpace(request.OrigenDesde) ? (object)DBNull.Value : request.OrigenDesde.Trim());
        command.Parameters.AddWithValue("@OrigenHasta", string.IsNullOrWhiteSpace(request.OrigenHasta) ? (object)DBNull.Value : request.OrigenHasta.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new LibroDiarioItemDto
            {
                Modo = reader.IsDBNull(reader.GetOrdinal("Modo")) ? string.Empty : reader.GetString(reader.GetOrdinal("Modo")),
                CodigoOrigen = reader.IsDBNull(reader.GetOrdinal("CodigoOrigen")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.IsDBNull(reader.GetOrdinal("NombreOrigen")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreOrigen")),
                Periodo = reader.IsDBNull(reader.GetOrdinal("Periodo")) ? string.Empty : reader.GetString(reader.GetOrdinal("Periodo")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? 0 : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                Item = reader.IsDBNull(reader.GetOrdinal("Item")) ? (short)0 : reader.GetInt16(reader.GetOrdinal("Item")),
                FechaEmision = reader.IsDBNull(reader.GetOrdinal("FechaEmision")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                CodigoCuenta = reader.IsDBNull(reader.GetOrdinal("CodigoCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.IsDBNull(reader.GetOrdinal("NombreCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreCuenta")),
                NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreAuxiliar = reader.IsDBNull(reader.GetOrdinal("NombreAuxiliar")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreAuxiliar")),
                TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? string.Empty : reader.GetString(reader.GetOrdinal("Serie")),
                Referencia = reader.IsDBNull(reader.GetOrdinal("Referencia")) ? string.Empty : reader.GetString(reader.GetOrdinal("Referencia")),
                Glosa = reader.IsDBNull(reader.GetOrdinal("Glosa")) ? string.Empty : reader.GetString(reader.GetOrdinal("Glosa")),
                TipoCambio = reader.IsDBNull(reader.GetOrdinal("TipoCambio")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                Debe = reader.IsDBNull(reader.GetOrdinal("Debe")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.IsDBNull(reader.GetOrdinal("Haber")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Haber")),
                DebeDolares = reader.IsDBNull(reader.GetOrdinal("DebeDolares")) ? 0m : reader.GetDecimal(reader.GetOrdinal("DebeDolares")),
                HaberDolares = reader.IsDBNull(reader.GetOrdinal("HaberDolares")) ? 0m : reader.GetDecimal(reader.GetOrdinal("HaberDolares"))
            });
        }

        return result;
    }
}
