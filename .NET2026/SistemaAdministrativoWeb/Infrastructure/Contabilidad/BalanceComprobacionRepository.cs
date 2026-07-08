using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BalanceComprobacionRepository(IDbConnectionFactory connectionFactory) : IBalanceComprobacionRepository
{
    public async Task<IReadOnlyCollection<BalanceComprobacionItemDto>> ListarAsync(BalanceComprobacionRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<BalanceComprobacionItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ReporteBalanceComprobacion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Anio", request.Anio);
        command.Parameters.AddWithValue("@PeriodoDesde", request.PeriodoDesde);
        command.Parameters.AddWithValue("@PeriodoHasta", request.PeriodoHasta);
        command.Parameters.AddWithValue("@Moneda", string.IsNullOrWhiteSpace(request.Moneda) ? "PEN" : request.Moneda.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@Grado", request.Grado);
        command.Parameters.AddWithValue("@TodasLasCuentas", request.TodasLasCuentas);
        command.Parameters.AddWithValue("@CuentaDesde", string.IsNullOrWhiteSpace(request.CuentaDesde) ? (object)DBNull.Value : request.CuentaDesde.Trim());
        command.Parameters.AddWithValue("@CuentaHasta", string.IsNullOrWhiteSpace(request.CuentaHasta) ? (object)DBNull.Value : request.CuentaHasta.Trim());
        command.Parameters.AddWithValue("@FiltrarGrado", request.FiltrarGrado);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new BalanceComprobacionItemDto
            {
                CodigoCuenta = reader.IsDBNull(reader.GetOrdinal("CodigoCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.IsDBNull(reader.GetOrdinal("NombreCuenta")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreCuenta")),
                ColBalance = reader.IsDBNull(reader.GetOrdinal("ColBalance")) ? string.Empty : reader.GetString(reader.GetOrdinal("ColBalance")),
                GradoCuenta = reader.IsDBNull(reader.GetOrdinal("GradoCuenta")) ? (byte)0 : Convert.ToByte(reader.GetValue(reader.GetOrdinal("GradoCuenta"))),
                DebAnt = reader.IsDBNull(reader.GetOrdinal("DebAnt")) ? 0m : reader.GetDecimal(reader.GetOrdinal("DebAnt")),
                HabAnt = reader.IsDBNull(reader.GetOrdinal("HabAnt")) ? 0m : reader.GetDecimal(reader.GetOrdinal("HabAnt")),
                DebMes = reader.IsDBNull(reader.GetOrdinal("DebMes")) ? 0m : reader.GetDecimal(reader.GetOrdinal("DebMes")),
                HabMes = reader.IsDBNull(reader.GetOrdinal("HabMes")) ? 0m : reader.GetDecimal(reader.GetOrdinal("HabMes")),
                Debe = reader.IsDBNull(reader.GetOrdinal("Debe")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.IsDBNull(reader.GetOrdinal("Haber")) ? 0m : reader.GetDecimal(reader.GetOrdinal("Haber"))
            });
        }

        return result;
    }
}
