using System.Data;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class VentaRepository(IDbConnectionFactory connectionFactory) : IVentaRepository
{
    public async Task<IReadOnlyCollection<VentaResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default)
    {
        var result = new List<VentaResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_VEN_ListarVentasPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", string.IsNullOrWhiteSpace(periodo) ? (object)DBNull.Value : periodo);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new VentaResumenDto
            {
                IdVenta = reader.GetInt32(reader.GetOrdinal("IdVenta")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdCliente = reader.GetInt32(reader.GetOrdinal("IdCliente")),
                CodigoCliente = reader.GetString(reader.GetOrdinal("CodigoCliente")),
                NombreCliente = reader.GetString(reader.GetOrdinal("NombreCliente")),
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
                Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
                Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetString(reader.GetOrdinal("Estado"))
            });
        }

        return result;
    }

    public async Task<PagedResult<VentaResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<VentaResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_VEN_ListarVentasPorEmpresa", connection)
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

        return new PagedResult<VentaResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<VentaDto?> ObtenerAsync(int idVenta, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_VEN_ObtenerVenta", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdVenta", idVenta);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        VentaDto? venta = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            venta = new VentaDto
            {
                IdVenta = reader.GetInt32(reader.GetOrdinal("IdVenta")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdCliente = reader.GetInt32(reader.GetOrdinal("IdCliente")),
                CodigoCliente = reader.GetString(reader.GetOrdinal("CodigoCliente")),
                TipoDocumentoCliente = reader.GetString(reader.GetOrdinal("TipoDocumento")),
                NumeroDocumentoCliente = reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                NombreCliente = reader.GetString(reader.GetOrdinal("NombreCliente")),
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                DescripcionConfiguracion = reader.GetString(reader.GetOrdinal("DescripcionConfiguracion")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
                Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
                Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
                OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
                Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                Estado = reader.GetString(reader.GetOrdinal("Estado"))
            };
        }

        if (venta is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                venta.Detalles.Add(new VentaDetalleDto
                {
                    IdVentaDetalle = reader.GetInt32(reader.GetOrdinal("IdVentaDetalle")),
                    IdVenta = reader.GetInt32(reader.GetOrdinal("IdVenta")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                    Cantidad = reader.GetDecimal(reader.GetOrdinal("Cantidad")),
                    ValorUnitario = reader.GetDecimal(reader.GetOrdinal("ValorUnitario")),
                    ImporteBruto = reader.GetDecimal(reader.GetOrdinal("ImporteBruto"))
                });
            }
        }

        return venta;
    }

    public async Task<GuardarVentaResultDto> GuardarAsync(GuardarVentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_VEN_GuardarVentaConAsiento", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdVenta", (object?)request.IdVenta ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdCliente", request.IdCliente);
        command.Parameters.AddWithValue("@IdConfiguracionContabilizacion", request.IdConfiguracionContabilizacion);
        command.Parameters.AddWithValue("@FechaEmision", request.FechaEmision.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@FechaContabilizacion", request.FechaContabilizacion.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@TipoComprobante", request.TipoComprobante);
        command.Parameters.AddWithValue("@Serie", request.Serie);
        command.Parameters.AddWithValue("@Numero", request.Numero);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@BaseImponible", request.BaseImponible);
        command.Parameters.AddWithValue("@Igv", request.Igv);
        command.Parameters.AddWithValue("@Isc", request.Isc);
        command.Parameters.AddWithValue("@OtrosTributos", request.OtrosTributos);
        command.Parameters.AddWithValue("@Redondeo", request.Redondeo);
        command.Parameters.AddWithValue("@ImporteTotal", request.ImporteTotal);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@DetalleXml", ConstruirDetalleXml(request.Detalles));
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el resultado de la venta grabada.");
        }

        return new GuardarVentaResultDto
        {
            IdVenta = reader.GetInt32(reader.GetOrdinal("IdVenta")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
            Estado = reader.GetString(reader.GetOrdinal("Estado"))
        };
    }

    private static VentaResumenDto MapearResumen(SqlDataReader reader)
    {
        return new VentaResumenDto
        {
            IdVenta = reader.GetInt32(reader.GetOrdinal("IdVenta")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdCliente = reader.GetInt32(reader.GetOrdinal("IdCliente")),
            CodigoCliente = reader.GetString(reader.GetOrdinal("CodigoCliente")),
            NombreCliente = reader.GetString(reader.GetOrdinal("NombreCliente")),
            IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
            ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
            EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
            FechaContabilizacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContabilizacion"))),
            Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
            TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
            Serie = reader.GetString(reader.GetOrdinal("Serie")),
            Numero = reader.GetString(reader.GetOrdinal("Numero")),
            IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
            CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
            BaseImponible = reader.GetDecimal(reader.GetOrdinal("BaseImponible")),
            Igv = reader.GetDecimal(reader.GetOrdinal("Igv")),
            Isc = reader.GetDecimal(reader.GetOrdinal("Isc")),
            OtrosTributos = reader.GetDecimal(reader.GetOrdinal("OtrosTributos")),
            Redondeo = reader.GetDecimal(reader.GetOrdinal("Redondeo")),
            ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
            Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
            Estado = reader.GetString(reader.GetOrdinal("Estado"))
        };
    }

    private static string ConstruirDetalleXml(IReadOnlyCollection<GuardarVentaDetalleRequest> detalles)
    {
        var xml = new XElement("Detalles",
            detalles.Select(x => new XElement("Detalle",
                new XAttribute("Item", x.Item),
                new XAttribute("Descripcion", x.Descripcion),
                new XAttribute("Cantidad", x.Cantidad.ToString("0.####", CultureInfo.InvariantCulture)),
                new XAttribute("ValorUnitario", x.ValorUnitario.ToString("0.######", CultureInfo.InvariantCulture)),
                new XAttribute("ImporteBruto", x.ImporteBruto.ToString("0.00", CultureInfo.InvariantCulture)))));

        return xml.ToString(SaveOptions.DisableFormatting);
    }
}
