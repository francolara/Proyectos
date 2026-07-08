using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AjusteCuentaRepository(IDbConnectionFactory connectionFactory) : IAjusteCuentaRepository
{
    public async Task<AjusteCuentaProcesoDto?> ObtenerAsync(int idEmpresa, string periodo, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerAjusteCuentaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", (periodo ?? string.Empty).Trim());

        await connection.OpenAsync(cancellationToken);
        return await LeerProcesoAsync(command, cancellationToken);
    }

    public async Task<AjusteCuentaProcesoDto> GenerarAsync(GenerarAjusteCuentaProcesoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GenerarAjusteCuentaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo.Trim());
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        var proceso = await LeerProcesoAsync(command, cancellationToken);
        return proceso ?? throw new InvalidOperationException("No se pudo obtener el proceso de ajuste de cuentas generado.");
    }

    public async Task EliminarAsync(int idEmpresa, string periodo, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarAjusteCuentaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", (periodo ?? string.Empty).Trim());
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(usuarioRegistro) ? (object)DBNull.Value : usuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<AjusteCuentaProcesoDto?> LeerProcesoAsync(SqlCommand command, CancellationToken cancellationToken)
    {
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        AjusteCuentaProcesoDto? proceso = null;
        if (await reader.ReadAsync(cancellationToken))
        {
            proceso = new AjusteCuentaProcesoDto
            {
                IdAjusteCuentaProceso = reader.GetInt32(reader.GetOrdinal("IdAjusteCuentaProceso")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                FechaAsiento = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaAsiento"))),
                TotalCuentas = reader.GetInt32(reader.GetOrdinal("TotalCuentas")),
                TotalAsientos = reader.GetInt32(reader.GetOrdinal("TotalAsientos")),
                TotalDebe = reader.GetDecimal(reader.GetOrdinal("TotalDebe")),
                TotalHaber = reader.GetDecimal(reader.GetOrdinal("TotalHaber")),
                FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
                UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
            };
        }

        if (proceso is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                proceso.Detalles.Add(new AjusteCuentaProcesoDetalleDto
                {
                    IdAjusteCuentaProcesoDetalle = reader.GetInt32(reader.GetOrdinal("IdAjusteCuentaProcesoDetalle")),
                    IdAjusteCuentaProceso = reader.GetInt32(reader.GetOrdinal("IdAjusteCuentaProceso")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                    TipoCambioAplicado = reader.GetDecimal(reader.GetOrdinal("TipoCambioAplicado")),
                    TotalAnalisis = reader.GetInt32(reader.GetOrdinal("TotalAnalisis")),
                    IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                    NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                    TotalDebe = reader.GetDecimal(reader.GetOrdinal("TotalDebe")),
                    TotalHaber = reader.GetDecimal(reader.GetOrdinal("TotalHaber")),
                    Estado = reader.GetString(reader.GetOrdinal("Estado")),
                    Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                    FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
                    UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
                });
            }
        }

        return proceso;
    }
}
