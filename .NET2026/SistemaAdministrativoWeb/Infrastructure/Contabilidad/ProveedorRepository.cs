using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ProveedorRepository(IDbConnectionFactory connectionFactory) : IProveedorRepository
{
    public async Task<IReadOnlyCollection<ProveedorDto>> ListarActivosPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        var result = new List<ProveedorDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_ADM_ListarProveedoresActivosPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new ProveedorDto
            {
                IdProveedor = reader.GetInt32(reader.GetOrdinal("IdProveedor")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
                CodigoProveedor = reader.GetString(reader.GetOrdinal("CodigoProveedor")),
                TipoDocumento = reader.GetString(reader.GetOrdinal("TipoDocumento")),
                NumeroDocumento = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreCompleto = reader.GetString(reader.GetOrdinal("NombreCompleto")),
                CorreoElectronico = reader.IsDBNull(reader.GetOrdinal("CorreoElectronico")) ? null : reader.GetString(reader.GetOrdinal("CorreoElectronico")),
                Telefono = reader.IsDBNull(reader.GetOrdinal("Telefono")) ? null : reader.GetString(reader.GetOrdinal("Telefono")),
                Contacto = reader.IsDBNull(reader.GetOrdinal("Contacto")) ? null : reader.GetString(reader.GetOrdinal("Contacto")),
                CuentaDetraccion = reader.IsDBNull(reader.GetOrdinal("CuentaDetraccion")) ? null : reader.GetString(reader.GetOrdinal("CuentaDetraccion")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }
}
