using System.Data;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CompraRepository(IDbConnectionFactory connectionFactory) : ICompraRepository
{
    public async Task<IReadOnlyCollection<CompraResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default)
    {
        var result = new List<CompraResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_ListarComprasPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", string.IsNullOrWhiteSpace(periodo) ? (object)DBNull.Value : periodo);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CompraResumenDto
            {
                IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdProveedor = reader.GetInt32(reader.GetOrdinal("IdProveedor")),
                CodigoProveedor = reader.GetString(reader.GetOrdinal("CodigoProveedor")),
                NombreProveedor = reader.GetString(reader.GetOrdinal("NombreProveedor")),
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                DescripcionTipoComprobante = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
                TotalExonerado = reader.GetDecimal(reader.GetOrdinal("TotalExonerado")),
                TotalInafecto = reader.GetDecimal(reader.GetOrdinal("TotalInafecto")),
                Icbper = reader.GetDecimal(reader.GetOrdinal("Icbper")),
                Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
                Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Saldo = reader.GetDecimal(reader.GetOrdinal("Saldo")),
                TieneDetraccion = reader.GetBoolean(reader.GetOrdinal("TieneDetraccion")),
                IdDetraccionSunat = reader.IsDBNull(reader.GetOrdinal("IdDetraccionSunat")) ? null : reader.GetInt32(reader.GetOrdinal("IdDetraccionSunat")),
                PorcentajeDetraccion = reader.GetDecimal(reader.GetOrdinal("PorcentajeDetraccion")),
                ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetString(reader.GetOrdinal("Estado")),
                Situacion = reader.GetString(reader.GetOrdinal("Situacion"))
            });
        }

        return result;
    }

    public async Task<IReadOnlyCollection<CompraDetraccionPendienteDto>> ListarDetraccionesPendientesPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default)
    {
        var result = new List<CompraDetraccionPendienteDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = connection.CreateCommand();
        command.CommandType = CommandType.Text;
        command.CommandText = """
            SELECT
                cd.IdCompraDetraccion,
                cd.IdCompra,
                cd.IdProveedor,
                p.CodigoProveedor,
                per.NombreCompleto AS NombreProveedor,
                per.NumeroDocumento AS NumeroDocumentoPersona,
                cd.FechaEmision,
                cd.FechaContabilizacion,
                cd.IdMoneda,
                m.CodigoMoneda,
                cd.TipoCambio,
                cd.CodigoDetraccionSunat,
                cd.DescripcionDetraccion,
                cd.PorcentajeDetraccion,
                cd.ImporteDetraccion,
                cd.Saldo,
                c.Serie,
                c.Numero,
                cd.ReferenciaDocumento,
                cd.Estado
            FROM dbo.COM_CompraDetraccion AS cd
            INNER JOIN dbo.COM_Compra AS c
                ON c.IdCompra = cd.IdCompra
               AND c.IdEmpresa = cd.IdEmpresa
            INNER JOIN dbo.ADM_Proveedor AS p
                ON p.IdProveedor = cd.IdProveedor
               AND p.IdEmpresa = cd.IdEmpresa
            INNER JOIN dbo.ADM_Persona AS per
                ON per.IdPersona = p.IdPersona
               AND per.IdEmpresa = cd.IdEmpresa
            INNER JOIN dbo.ADM_Moneda AS m
                ON m.IdMoneda = cd.IdMoneda
            WHERE cd.IdEmpresa = @IdEmpresa
              AND cd.Saldo > 0
              AND (@Periodo IS NULL
                   OR (CONVERT(CHAR(4), YEAR(cd.FechaContabilizacion))
                       + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(cd.FechaContabilizacion)), 2)) = @Periodo)
            ORDER BY
                cd.FechaEmision DESC,
                cd.IdCompraDetraccion DESC;
            """;

        command.Parameters.Add(new SqlParameter("@IdEmpresa", SqlDbType.Int) { Value = idEmpresa });
        command.Parameters.Add(new SqlParameter("@Periodo", SqlDbType.Char, 6)
        {
            Value = string.IsNullOrWhiteSpace(periodo) ? DBNull.Value : periodo.Trim()
        });

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CompraDetraccionPendienteDto
            {
                IdCompraDetraccion = reader.GetInt32(reader.GetOrdinal("IdCompraDetraccion")),
                IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
                IdProveedor = reader.GetInt32(reader.GetOrdinal("IdProveedor")),
                CodigoProveedor = reader.GetString(reader.GetOrdinal("CodigoProveedor")),
                NombreProveedor = reader.GetString(reader.GetOrdinal("NombreProveedor")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                CodigoDetraccionSunat = reader.GetString(reader.GetOrdinal("CodigoDetraccionSunat")),
                DescripcionDetraccion = reader.GetString(reader.GetOrdinal("DescripcionDetraccion")),
                PorcentajeDetraccion = reader.GetDecimal(reader.GetOrdinal("PorcentajeDetraccion")),
                ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
                Saldo = reader.GetDecimal(reader.GetOrdinal("Saldo")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                ReferenciaDocumento = reader.GetString(reader.GetOrdinal("ReferenciaDocumento")),
                Estado = reader.GetString(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }

    public async Task<PagedResult<CompraResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<CompraResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_ListarComprasPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Ejercicio", ejercicio);
        command.Parameters.AddWithValue("@Mes", mes);
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

        return new PagedResult<CompraResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<CompraDto?> ObtenerAsync(int idCompra, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_ObtenerCompra", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCompra", idCompra);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        CompraDto? compra = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            compra = new CompraDto
            {
                IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdProveedor = reader.GetInt32(reader.GetOrdinal("IdProveedor")),
                CodigoProveedor = reader.GetString(reader.GetOrdinal("CodigoProveedor")),
                TipoDocumentoProveedor = reader.GetString(reader.GetOrdinal("TipoDocumento")),
                NumeroDocumentoProveedor = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreProveedor = reader.GetString(reader.GetOrdinal("NombreProveedor")),
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                DescripcionConfiguracion = reader.GetString(reader.GetOrdinal("DescripcionConfiguracion")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                DescripcionTipoComprobante = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
                TotalExonerado = reader.GetDecimal(reader.GetOrdinal("TotalExonerado")),
                TotalInafecto = reader.GetDecimal(reader.GetOrdinal("TotalInafecto")),
                Icbper = reader.GetDecimal(reader.GetOrdinal("Icbper")),
                Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
                Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Saldo = reader.GetDecimal(reader.GetOrdinal("Saldo")),
                IdCompraDetraccion = reader.IsDBNull(reader.GetOrdinal("IdCompraDetraccion")) ? null : reader.GetInt32(reader.GetOrdinal("IdCompraDetraccion")),
                IdAsientoDetraccion = reader.IsDBNull(reader.GetOrdinal("IdAsientoDetraccion")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsientoDetraccion")),
                TieneDetraccion = reader.GetBoolean(reader.GetOrdinal("TieneDetraccion")),
                IdDetraccionSunat = reader.IsDBNull(reader.GetOrdinal("IdDetraccionSunat")) ? null : reader.GetInt32(reader.GetOrdinal("IdDetraccionSunat")),
                CodigoDetraccionSunat = reader.IsDBNull(reader.GetOrdinal("CodigoDetraccionSunat")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoDetraccionSunat")),
                DescripcionDetraccionSunat = reader.IsDBNull(reader.GetOrdinal("DescripcionDetraccionSunat")) ? string.Empty : reader.GetString(reader.GetOrdinal("DescripcionDetraccionSunat")),
                PorcentajeDetraccion = reader.GetDecimal(reader.GetOrdinal("PorcentajeDetraccion")),
                ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetString(reader.GetOrdinal("Estado"))
            };
        }

        if (compra is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                compra.Detalles.Add(new CompraDetalleDto
                {
                    IdCompraDetalle = reader.GetInt32(reader.GetOrdinal("IdCompraDetalle")),
                    IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    IdTipoAfectacionIGV = reader.GetInt32(reader.GetOrdinal("IdTipoAfectacionIGV")),
                    CodigoAfectacionIGV = reader.GetString(reader.GetOrdinal("CodigoAfectacionIGV")),
                    NombreAfectacionIGV = reader.GetString(reader.GetOrdinal("NombreAfectacionIGV")),
                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                    Cantidad = reader.GetDecimal(reader.GetOrdinal("Cantidad")),
                    ValorUnitario = reader.GetDecimal(reader.GetOrdinal("ValorUnitario")),
                    ImporteBruto = reader.GetDecimal(reader.GetOrdinal("ImporteBruto"))
                });
            }
        }

        return compra;
    }

    public async Task<GuardarCompraResultDto> GuardarAsync(GuardarCompraRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_GuardarCompraConAsiento", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCompra", (object?)request.IdCompra ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdProveedor", request.IdProveedor);
        command.Parameters.AddWithValue("@IdConfiguracionContabilizacion", request.IdConfiguracionContabilizacion);
        command.Parameters.AddWithValue("@FechaEmision", request.FechaEmision.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@FechaContabilizacion", request.FechaContabilizacion.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@TipoComprobante", request.TipoComprobante);
        command.Parameters.AddWithValue("@Serie", request.Serie);
        command.Parameters.AddWithValue("@Numero", request.Numero);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@BaseImponible", request.BaseImponible);
        command.Parameters.AddWithValue("@TotalExonerado", request.TotalExonerado);
        command.Parameters.AddWithValue("@TotalInafecto", request.TotalInafecto);
        command.Parameters.AddWithValue("@Icbper", request.Icbper);
        command.Parameters.AddWithValue("@Igv", request.Igv);
        command.Parameters.AddWithValue("@Isc", request.Isc);
        command.Parameters.AddWithValue("@OtrosTributos", request.OtrosTributos);
        command.Parameters.AddWithValue("@Redondeo", request.Redondeo);
        command.Parameters.AddWithValue("@ImporteTotal", request.ImporteTotal);
        command.Parameters.AddWithValue("@TieneDetraccion", request.TieneDetraccion);
        command.Parameters.AddWithValue("@IdDetraccionSunat", (object?)request.IdDetraccionSunat ?? DBNull.Value);
        command.Parameters.AddWithValue("@ImporteDetraccion", request.ImporteDetraccion);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@DetalleXml", ConstruirDetalleXml(request.Detalles));
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el resultado de la compra grabada.");
        }

        return new GuardarCompraResultDto
        {
            IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            IdAsientoDetraccion = reader.IsDBNull(reader.GetOrdinal("IdAsientoDetraccion")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsientoDetraccion")),
            ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
            ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
            Estado = reader.GetString(reader.GetOrdinal("Estado"))
        };
    }

    public async Task EliminarAsync(int idCompra, int idEmpresa, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_COM_EliminarCompra", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCompra", idCompra);
        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static CompraResumenDto MapearResumen(SqlDataReader reader)
    {
        return new CompraResumenDto
        {
            IdCompra = reader.GetInt32(reader.GetOrdinal("IdCompra")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdProveedor = reader.GetInt32(reader.GetOrdinal("IdProveedor")),
            CodigoProveedor = reader.GetString(reader.GetOrdinal("CodigoProveedor")),
            NombreProveedor = reader.GetString(reader.GetOrdinal("NombreProveedor")),
            IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
            ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
            EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
            FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
            Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
            TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
            DescripcionTipoComprobante = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobante")),
            Serie = reader.GetString(reader.GetOrdinal("Serie")),
            Numero = reader.GetString(reader.GetOrdinal("Numero")),
            NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
            IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
            CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
            BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
            TotalExonerado = reader.GetDecimal(reader.GetOrdinal("TotalExonerado")),
            TotalInafecto = reader.GetDecimal(reader.GetOrdinal("TotalInafecto")),
            Icbper = reader.GetDecimal(reader.GetOrdinal("Icbper")),
            Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
            Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
            OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
            Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
            ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
            Saldo = reader.GetDecimal(reader.GetOrdinal("Saldo")),
            TieneDetraccion = reader.GetBoolean(reader.GetOrdinal("TieneDetraccion")),
            IdDetraccionSunat = reader.IsDBNull(reader.GetOrdinal("IdDetraccionSunat")) ? null : reader.GetInt32(reader.GetOrdinal("IdDetraccionSunat")),
            PorcentajeDetraccion = reader.GetDecimal(reader.GetOrdinal("PorcentajeDetraccion")),
            ImporteDetraccion = reader.GetDecimal(reader.GetOrdinal("ImporteDetraccion")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
            Estado = reader.GetString(reader.GetOrdinal("Estado")),
            Situacion = reader.GetString(reader.GetOrdinal("Situacion"))
        };
    }

    private static string ConstruirDetalleXml(IReadOnlyCollection<GuardarCompraDetalleRequest> detalles)
    {
        var xml = new XElement("Detalles",
            detalles.Select(x => new XElement("Detalle",
                new XAttribute("Item", x.Item),
                new XAttribute("IdPlanCuenta", x.IdPlanCuenta),
                new XAttribute("IdTipoAfectacionIGV", x.IdTipoAfectacionIGV),
                new XAttribute("Descripcion", x.Descripcion),
                new XAttribute("Cantidad", x.Cantidad.ToString("0.####", CultureInfo.InvariantCulture)),
                new XAttribute("ValorUnitario", x.ValorUnitario.ToString("0.######", CultureInfo.InvariantCulture)),
                new XAttribute("ImporteBruto", x.ImporteBruto.ToString("0.00", CultureInfo.InvariantCulture)))));

        return xml.ToString(SaveOptions.DisableFormatting);
    }
}
