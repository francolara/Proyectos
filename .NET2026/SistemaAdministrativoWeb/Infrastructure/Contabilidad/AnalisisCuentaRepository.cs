using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AnalisisCuentaRepository(IDbConnectionFactory connectionFactory) : IAnalisisCuentaRepository
{
    public async Task<IReadOnlyCollection<AnalisisCuentaItemDto>> ListarAsync(AnalisisCuentaRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<AnalisisCuentaItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ReporteAnalisisCuentas", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo);
        command.Parameters.AddWithValue("@CuentaDesde", string.IsNullOrWhiteSpace(request.CuentaDesde) ? (object)DBNull.Value : request.CuentaDesde.Trim());
        command.Parameters.AddWithValue("@CuentaHasta", string.IsNullOrWhiteSpace(request.CuentaHasta) ? (object)DBNull.Value : request.CuentaHasta.Trim());
        command.Parameters.AddWithValue("@Auxiliar", string.IsNullOrWhiteSpace(request.Auxiliar) ? (object)DBNull.Value : request.Auxiliar.Trim());
        command.Parameters.AddWithValue("@Moneda", string.IsNullOrWhiteSpace(request.Moneda) ? "PEN" : request.Moneda.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@Estado", string.IsNullOrWhiteSpace(request.Estado) ? "T" : request.Estado.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@Tipo", string.IsNullOrWhiteSpace(request.Tipo) ? "0" : request.Tipo.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new AnalisisCuentaItemDto
            {
                CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                Auxiliar = reader.IsDBNull(reader.GetOrdinal("Auxiliar")) ? string.Empty : reader.GetString(reader.GetOrdinal("Auxiliar")),
                NombreAuxiliar = reader.IsDBNull(reader.GetOrdinal("NombreAuxiliar")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreAuxiliar")),
                TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? string.Empty : reader.GetString(reader.GetOrdinal("Serie")),
                NumeroReferencia = reader.IsDBNull(reader.GetOrdinal("NumeroReferencia")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroReferencia")),
                Periodo = reader.IsDBNull(reader.GetOrdinal("Periodo")) ? string.Empty : reader.GetString(reader.GetOrdinal("Periodo")),
                Comprobante = reader.IsDBNull(reader.GetOrdinal("Comprobante")) ? string.Empty : reader.GetString(reader.GetOrdinal("Comprobante")),
                GlosaDetalle = reader.IsDBNull(reader.GetOrdinal("GlosaDetalle")) ? string.Empty : reader.GetString(reader.GetOrdinal("GlosaDetalle")),
                FechaEmision = reader.IsDBNull(reader.GetOrdinal("FechaEmision")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
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
