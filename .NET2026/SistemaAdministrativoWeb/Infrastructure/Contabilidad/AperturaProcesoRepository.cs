using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AperturaProcesoRepository(IDbConnectionFactory connectionFactory) : IAperturaProcesoRepository
{
    public async Task<AperturaProcesoDto?> ObtenerAsync(int idEmpresa, short anioApertura, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerAperturaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@AnioApertura", anioApertura);

        await connection.OpenAsync(cancellationToken);
        return await LeerProcesoAsync(command, cancellationToken);
    }

    public async Task<AperturaProcesoDto> GenerarAsync(GenerarAperturaProcesoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GenerarAperturaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@AnioApertura", request.AnioApertura);
        command.Parameters.AddWithValue("@MesSaldoHasta", request.MesSaldoHasta);
        command.Parameters.AddWithValue("@TipoCambioCompra", request.TipoCambioCompra);
        command.Parameters.AddWithValue("@TipoCambioVenta", request.TipoCambioVenta);
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        var proceso = await LeerProcesoAsync(command, cancellationToken);
        return proceso ?? throw new InvalidOperationException("No se pudo obtener el proceso de apertura generado.");
    }

    public async Task EliminarAsync(int idEmpresa, short anioApertura, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarAperturaProceso", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@AnioApertura", anioApertura);
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(usuarioRegistro) ? (object)DBNull.Value : usuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<AperturaProcesoDto?> LeerProcesoAsync(SqlCommand command, CancellationToken cancellationToken)
    {
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        AperturaProcesoDto? proceso = null;
        if (await reader.ReadAsync(cancellationToken))
        {
            proceso = new AperturaProcesoDto
            {
                IdAperturaProceso = reader.GetInt32(reader.GetOrdinal("IdAperturaProceso")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                AnioApertura = reader.GetInt16(reader.GetOrdinal("AnioApertura")),
                AnioSaldo = reader.GetInt16(reader.GetOrdinal("AnioSaldo")),
                MesSaldoHasta = reader.GetByte(reader.GetOrdinal("MesSaldoHasta")),
                PeriodoSaldoHasta = reader.GetString(reader.GetOrdinal("PeriodoSaldoHasta")),
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                FechaAsiento = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaAsiento"))),
                UsaTipoCambioSbs = reader.GetBoolean(reader.GetOrdinal("UsaTipoCambioSbs")),
                TipoCambioCompra = reader.GetDecimal(reader.GetOrdinal("TipoCambioCompra")),
                TipoCambioVenta = reader.GetDecimal(reader.GetOrdinal("TipoCambioVenta")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                TotalLineas = reader.GetInt32(reader.GetOrdinal("TotalLineas")),
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
                proceso.Detalles.Add(new AperturaProcesoDetalleDto
                {
                    IdAperturaProcesoDetalle = reader.GetInt32(reader.GetOrdinal("IdAperturaProcesoDetalle")),
                    IdAperturaProceso = reader.GetInt32(reader.GetOrdinal("IdAperturaProceso")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    TipoDetalle = reader.GetString(reader.GetOrdinal("TipoDetalle")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                    TipoCambioAplicado = reader.GetDecimal(reader.GetOrdinal("TipoCambioAplicado")),
                    TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? null : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                    Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? null : reader.GetString(reader.GetOrdinal("Serie")),
                    NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? null : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                    Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                    Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                    TotalImporteS = reader.GetDecimal(reader.GetOrdinal("TotalImporteS")),
                    TotalImporteD = reader.GetDecimal(reader.GetOrdinal("TotalImporteD")),
                    Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                    FechaRegistro = reader.GetDateTime(reader.GetOrdinal("FechaRegistro")),
                    UsuarioRegistro = reader.IsDBNull(reader.GetOrdinal("UsuarioRegistro")) ? null : reader.GetString(reader.GetOrdinal("UsuarioRegistro"))
                });
            }
        }

        return proceso;
    }
}
