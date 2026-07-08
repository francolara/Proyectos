using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroMayorRepository(IDbConnectionFactory connectionFactory) : ILibroMayorRepository
{
    public async Task<IReadOnlyCollection<LibroMayorItemDto>> ListarAsync(LibroMayorRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<LibroMayorItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ReporteLibroMayor", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo);
        command.Parameters.AddWithValue("@CuentaDesde", string.IsNullOrWhiteSpace(request.CuentaDesde) ? (object)DBNull.Value : request.CuentaDesde.Trim());
        command.Parameters.AddWithValue("@CuentaHasta", string.IsNullOrWhiteSpace(request.CuentaHasta) ? (object)DBNull.Value : request.CuentaHasta.Trim());
        command.Parameters.AddWithValue("@NumeroDocumento", string.IsNullOrWhiteSpace(request.NumeroDocumento) ? (object)DBNull.Value : request.NumeroDocumento.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new LibroMayorItemDto
            {
                CodigoCuenta = reader.IsDBNull(reader.GetOrdinal("CodigoCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.IsDBNull(reader.GetOrdinal("NombreCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreCuenta")),
                CodigoOrigen = reader.IsDBNull(reader.GetOrdinal("CodigoOrigen")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.IsDBNull(reader.GetOrdinal("NombreOrigen")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreOrigen")),
                Periodo = reader.IsDBNull(reader.GetOrdinal("Periodo")) ? string.Empty : reader.GetString(reader.GetOrdinal("Periodo")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? 0 : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                Item = reader.IsDBNull(reader.GetOrdinal("Item")) ? (short)0 : reader.GetInt16(reader.GetOrdinal("Item")),
                FechaEmision = reader.IsDBNull(reader.GetOrdinal("FechaEmision")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? string.Empty : reader.GetString(reader.GetOrdinal("Serie")),
                Referencia = reader.IsDBNull(reader.GetOrdinal("Referencia")) ? string.Empty : reader.GetString(reader.GetOrdinal("Referencia")),
                NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreAuxiliar = reader.IsDBNull(reader.GetOrdinal("NombreAuxiliar")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreAuxiliar")),
                Glosa = reader.IsDBNull(reader.GetOrdinal("Glosa")) ? string.Empty : reader.GetString(reader.GetOrdinal("Glosa")),
                TipoCambio = reader.IsDBNull(reader.GetOrdinal("TipoCambio")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                Debe = reader.IsDBNull(reader.GetOrdinal("Debe")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.IsDBNull(reader.GetOrdinal("Haber")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Haber")),
                DebeDolares = reader.IsDBNull(reader.GetOrdinal("DebeDolares")) ? 0m : reader.GetDecimal(reader.GetOrdinal("DebeDolares")),
                HaberDolares = reader.IsDBNull(reader.GetOrdinal("HaberDolares")) ? 0m : reader.GetDecimal(reader.GetOrdinal("HaberDolares")),
                SaldoInicial = reader.IsDBNull(reader.GetOrdinal("SaldoInicial")) ? 0m : reader.GetDecimal(reader.GetOrdinal("SaldoInicial")),
                SaldoInicialDolares = reader.IsDBNull(reader.GetOrdinal("SaldoInicialDolares")) ? 0m : reader.GetDecimal(reader.GetOrdinal("SaldoInicialDolares")),
                EsSaldoInicial = !reader.IsDBNull(reader.GetOrdinal("EsSaldoInicial")) && reader.GetBoolean(reader.GetOrdinal("EsSaldoInicial"))
            });
        }

        return result;
    }
}
