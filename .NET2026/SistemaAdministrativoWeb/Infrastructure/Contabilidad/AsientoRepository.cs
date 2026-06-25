using System.Data;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoRepository(IDbConnectionFactory connectionFactory) : IAsientoRepository
{
    public async Task<IReadOnlyCollection<AsientoResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, bool soloManual = false, CancellationToken cancellationToken = default)
    {
        var result = new List<AsientoResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarAsientosPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", string.IsNullOrWhiteSpace(periodo) ? (object)DBNull.Value : periodo);
        command.Parameters.AddWithValue("@SoloManual", soloManual);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(MapearResumen(reader));
        }

        return result;
    }

    public async Task<PagedResult<AsientoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloManual = false, CancellationToken cancellationToken = default)
    {
        var result = new List<AsientoResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarAsientosPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Ejercicio", ejercicio);
        command.Parameters.AddWithValue("@Mes", mes);
        command.Parameters.AddWithValue("@SoloManual", soloManual);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(MapearResumen(reader));
        }

        return new PagedResult<AsientoResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<AsientoDto?> ObtenerAsync(int idAsiento, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerAsiento", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdAsiento", idAsiento);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        AsientoDto? asiento = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            var resumen = MapearResumen(reader);
            asiento = new AsientoDto
            {
                IdAsiento = resumen.IdAsiento,
                IdEmpresa = resumen.IdEmpresa,
                IdOrigen = resumen.IdOrigen,
                CodigoOrigen = resumen.CodigoOrigen,
                NombreOrigen = resumen.NombreOrigen,
                PermiteRegistroManual = resumen.PermiteRegistroManual,
                Ejercicio = resumen.Ejercicio,
                Mes = resumen.Mes,
                Periodo = resumen.Periodo,
                NumeroAsiento = resumen.NumeroAsiento,
                FechaEmision = resumen.FechaEmision,
                FechaAsiento = resumen.FechaAsiento,
                Glosa = resumen.Glosa,
                IdMoneda = resumen.IdMoneda,
                CodigoMoneda = resumen.CodigoMoneda,
                NombreMoneda = resumen.NombreMoneda,
                SimboloMoneda = resumen.SimboloMoneda,
                TipoCambio = resumen.TipoCambio,
                TotalDebe = resumen.TotalDebe,
                TotalHaber = resumen.TotalHaber,
                Estado = resumen.Estado,
                ReferenciaExterna = resumen.ReferenciaExterna,
                Observacion = resumen.Observacion
            };
        }

        if (asiento is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                asiento.Detalles.Add(new AsientoDetalleDto
                {
                    IdAsientoDetalle = reader.GetInt32(reader.GetOrdinal("IdAsientoDetalle")),
                    IdAsiento = reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    GlosaDetalle = reader.IsDBNull(reader.GetOrdinal("GlosaDetalle")) ? null : reader.GetString(reader.GetOrdinal("GlosaDetalle")),
                    CodigoCentroCosto = reader.IsDBNull(reader.GetOrdinal("CodigoCentroCosto")) ? null : reader.GetString(reader.GetOrdinal("CodigoCentroCosto")),
                    TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? null : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                    NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? null : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                    Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? null : reader.GetString(reader.GetOrdinal("Serie")),
                    TipoCambioLinea = reader.IsDBNull(reader.GetOrdinal("TipoCambioLinea")) ? null : reader.GetDecimal(reader.GetOrdinal("TipoCambioLinea")),
                    Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                    Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                    ReferenciaLinea = reader.IsDBNull(reader.GetOrdinal("ReferenciaLinea")) ? null : reader.GetString(reader.GetOrdinal("ReferenciaLinea"))
                });
            }
        }

        return asiento;
    }

    public async Task<GuardarAsientoManualResultDto> GuardarManualAsync(GuardarAsientoManualRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarAsientoManual", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdAsiento", (object?)request.IdAsiento ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdOrigen", request.IdOrigen);
        command.Parameters.AddWithValue("@FechaEmision", request.FechaEmision.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@FechaAsiento", request.FechaAsiento.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@Glosa", request.Glosa);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@ReferenciaExterna", (object?)request.ReferenciaExterna ?? DBNull.Value);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@DetalleXml", ConstruirDetalleXml(request.Detalles));
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo recuperar el resultado del asiento grabado.");
        }

        return new GuardarAsientoManualResultDto
        {
            IdAsiento = reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
            NumeroAsiento = reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
            TotalDebe = reader.GetDecimal(reader.GetOrdinal("TotalDebe")),
            TotalHaber = reader.GetDecimal(reader.GetOrdinal("TotalHaber")),
            Estado = reader.GetString(reader.GetOrdinal("Estado"))
        };
    }

    public async Task EliminarAsync(int idAsiento, int idEmpresa, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarAsiento", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdAsiento", idAsiento);
        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static AsientoResumenDto MapearResumen(SqlDataReader reader)
    {
        return new AsientoResumenDto
        {
            IdAsiento = reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
            CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
            NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
            PermiteRegistroManual = reader.GetBoolean(reader.GetOrdinal("PermiteRegistroManual")),
            Ejercicio = reader.GetInt16(reader.GetOrdinal("Ejercicio")),
            Mes = reader.GetByte(reader.GetOrdinal("Mes")),
            Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
            NumeroAsiento = reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
            FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
            FechaAsiento = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaAsiento"))),
            Glosa = reader.GetString(reader.GetOrdinal("Glosa")),
            IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
            CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            NombreMoneda = reader.GetString(reader.GetOrdinal("NombreMoneda")),
            SimboloMoneda = reader.GetString(reader.GetOrdinal("SimboloMoneda")),
            TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
            TotalDebe = reader.GetDecimal(reader.GetOrdinal("TotalDebe")),
            TotalHaber = reader.GetDecimal(reader.GetOrdinal("TotalHaber")),
            Estado = reader.GetString(reader.GetOrdinal("Estado")),
            ReferenciaExterna = reader.IsDBNull(reader.GetOrdinal("ReferenciaExterna")) ? null : reader.GetString(reader.GetOrdinal("ReferenciaExterna")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion"))
        };
    }

    private static string ConstruirDetalleXml(IReadOnlyCollection<GuardarAsientoDetalleRequest> detalles)
    {
        var xml = new XElement("Detalles",
            detalles.Select(x => new XElement("Detalle",
                new XAttribute("Item", x.Item),
                new XAttribute("IdPlanCuenta", x.IdPlanCuenta),
                new XAttribute("GlosaDetalle", x.GlosaDetalle ?? string.Empty),
                new XAttribute("CodigoCentroCosto", x.CodigoCentroCosto ?? string.Empty),
                new XAttribute("TipoDocumento", x.TipoDocumento ?? string.Empty),
                new XAttribute("NumeroDocumento", x.NumeroDocumento ?? string.Empty),
                new XAttribute("Serie", x.Serie ?? string.Empty),
                new XAttribute("TipoCambioLinea", (x.TipoCambioLinea ?? 0m).ToString("0.000000", CultureInfo.InvariantCulture)),
                new XAttribute("Debe", x.Debe.ToString("0.00", CultureInfo.InvariantCulture)),
                new XAttribute("Haber", x.Haber.ToString("0.00", CultureInfo.InvariantCulture)),
                new XAttribute("ReferenciaLinea", x.ReferenciaLinea ?? string.Empty))));

        return xml.ToString(SaveOptions.DisableFormatting);
    }
}
