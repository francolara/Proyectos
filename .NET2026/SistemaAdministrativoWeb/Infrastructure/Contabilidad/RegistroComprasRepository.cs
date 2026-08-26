using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class RegistroComprasRepository(IDbConnectionFactory connectionFactory) : IRegistroComprasRepository
{
    public async Task<IReadOnlyCollection<RegistroComprasItemDto>> ListarAsync(RegistroComprasRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        var result = new List<RegistroComprasItemDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_ReporteRegistroCompras", connection)
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
            result.Add(new RegistroComprasItemDto
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
                BaseImponibleGravada = reader.GetDecimal(reader.GetOrdinal("BaseImponibleGravada")),
                IgvGravado = reader.GetDecimal(reader.GetOrdinal("IgvGravado")),
                BaseImponibleGasto = reader.GetDecimal(reader.GetOrdinal("BaseImponibleGasto")),
                IgvGasto = reader.GetDecimal(reader.GetOrdinal("IgvGasto")),
                BaseImponibleSinCredito = reader.GetDecimal(reader.GetOrdinal("BaseImponibleSinCredito")),
                IgvSinCredito = reader.GetDecimal(reader.GetOrdinal("IgvSinCredito")),
                TotalExonerado = reader.GetDecimal(reader.GetOrdinal("TotalExonerado")),
                TotalInafecto = reader.GetDecimal(reader.GetOrdinal("TotalInafecto")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Icbper = reader.GetDecimal(reader.GetOrdinal("Icbper")),
                Retencion = reader.GetDecimal(reader.GetOrdinal("Retencion")),
                ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
                ImportePercepcion = reader.GetDecimal(reader.GetOrdinal("ImportePercepcion")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Estado = reader.GetString(reader.GetOrdinal("Estado")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("Observacion"))
            });
        }

        return result;
    }
}
