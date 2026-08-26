using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class RegistroVentasRepository(IDbConnectionFactory connectionFactory) : IRegistroVentasRepository
{
    public async Task<IReadOnlyCollection<RegistroVentasItemDto>> ListarAsync(RegistroVentasRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<RegistroVentasItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_VEN_ReporteRegistroVentas", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Anio", request.Anio);
        command.Parameters.AddWithValue("@Mes", request.Mes);
        command.Parameters.AddWithValue("@CodigoPersona", string.IsNullOrWhiteSpace(request.CodigoPersona) ? (object)DBNull.Value : request.CodigoPersona.Trim());
        command.Parameters.AddWithValue("@NumeroDocumento", string.IsNullOrWhiteSpace(request.NumeroComprobante) ? (object)DBNull.Value : request.NumeroComprobante.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new RegistroVentasItemDto
            {
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                DescripcionTipoComprobante = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                CodigoPersona = reader.GetString(reader.GetOrdinal("CodigoPersona")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                NombrePersona = reader.GetString(reader.GetOrdinal("NombrePersona")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
                Descuento = reader.GetDecimal(reader.GetOrdinal("Descuento")),
                TotalExonerado = reader.GetDecimal(reader.GetOrdinal("TotalExonerado")),
                TotalInafecto = reader.GetDecimal(reader.GetOrdinal("TotalInafecto")),
                Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
                Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Icbper = reader.GetDecimal(reader.GetOrdinal("Icbper")),
                Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Estado = reader.GetString(reader.GetOrdinal("Estado")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("Observacion"))
            });
        }

        return result;
    }
}
