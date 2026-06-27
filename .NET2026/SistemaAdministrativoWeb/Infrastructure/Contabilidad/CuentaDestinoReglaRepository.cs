using System.Data;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CuentaDestinoReglaRepository(IDbConnectionFactory connectionFactory) : ICuentaDestinoReglaRepository
{
    public async Task<IReadOnlyCollection<CuentaDestinoReglaResumenDto>> ListarPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaDestinoReglaResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarCuentasDestinoReglaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new CuentaDestinoReglaResumenDto
            {
                IdCuentaDestinoRegla = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoRegla")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdPlanCuentaOrigen = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaOrigen")),
                CodigoCuentaOrigen = reader.GetString(reader.GetOrdinal("CodigoCuentaOrigen")),
                NombreCuentaOrigen = reader.GetString(reader.GetOrdinal("NombreCuentaOrigen")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                CantidadTramos = reader.GetInt32(reader.GetOrdinal("CantidadTramos")),
                PorcentajeTotal = reader.GetDecimal(reader.GetOrdinal("PorcentajeTotal"))
            });
        }

        return result;
    }

    public async Task<PagedResult<CuentaDestinoReglaResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<CuentaDestinoReglaResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarCuentasDestinoReglaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new CuentaDestinoReglaResumenDto
            {
                IdCuentaDestinoRegla = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoRegla")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdPlanCuentaOrigen = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaOrigen")),
                CodigoCuentaOrigen = reader.GetString(reader.GetOrdinal("CodigoCuentaOrigen")),
                NombreCuentaOrigen = reader.GetString(reader.GetOrdinal("NombreCuentaOrigen")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                CantidadTramos = reader.GetInt32(reader.GetOrdinal("CantidadTramos")),
                PorcentajeTotal = reader.GetDecimal(reader.GetOrdinal("PorcentajeTotal"))
            });
        }

        return new PagedResult<CuentaDestinoReglaResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<CuentaDestinoReglaDto?> ObtenerAsync(int idCuentaDestinoRegla, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerCuentaDestinoRegla", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaDestinoRegla", idCuentaDestinoRegla);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        CuentaDestinoReglaDto? regla = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            regla = new CuentaDestinoReglaDto
            {
                IdCuentaDestinoRegla = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoRegla")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdPlanCuentaOrigen = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaOrigen")),
                CodigoCuentaOrigen = reader.GetString(reader.GetOrdinal("CodigoCuentaOrigen")),
                NombreCuentaOrigen = reader.GetString(reader.GetOrdinal("NombreCuentaOrigen")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion"))
            };
        }

        if (regla is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                regla.Detalles.Add(new CuentaDestinoReglaDetalleDto
                {
                    IdCuentaDestinoReglaDetalle = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoReglaDetalle")),
                    IdCuentaDestinoRegla = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoRegla")),
                    Orden = reader.GetInt16(reader.GetOrdinal("Orden")),
                    IdPlanCuentaDestinoCargo = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaDestinoCargo")),
                    CodigoCuentaDestinoCargo = reader.GetString(reader.GetOrdinal("CodigoCuentaDestinoCargo")),
                    NombreCuentaDestinoCargo = reader.GetString(reader.GetOrdinal("NombreCuentaDestinoCargo")),
                    IdPlanCuentaDestinoAbono = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaDestinoAbono")),
                    CodigoCuentaDestinoAbono = reader.GetString(reader.GetOrdinal("CodigoCuentaDestinoAbono")),
                    NombreCuentaDestinoAbono = reader.GetString(reader.GetOrdinal("NombreCuentaDestinoAbono")),
                    Porcentaje = reader.GetDecimal(reader.GetOrdinal("Porcentaje")),
                    Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
                });
            }
        }

        return regla;
    }

    public async Task<int> GuardarAsync(GuardarCuentaDestinoReglaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarCuentaDestinoRegla", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdPlanCuentaOrigen", request.IdPlanCuentaOrigen);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@Observacion", (object?)request.Observacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@DetalleXml", ConstruirDetalleXml(request.Detalles));
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el identificador de la regla grabada.");
        }

        return reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoRegla"));
    }

    public async Task EliminarAsync(int idCuentaDestinoRegla, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarCuentaDestinoRegla", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdCuentaDestinoRegla", idCuentaDestinoRegla);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static string ConstruirDetalleXml(IReadOnlyCollection<GuardarCuentaDestinoReglaDetalleRequest> detalles)
    {
        var xml = new XElement("Detalles",
            detalles.Select(x => new XElement("Detalle",
                new XAttribute("Orden", x.Orden),
                new XAttribute("IdPlanCuentaDestinoCargo", x.IdPlanCuentaDestinoCargo),
                new XAttribute("IdPlanCuentaDestinoAbono", x.IdPlanCuentaDestinoAbono),
                new XAttribute("Porcentaje", x.Porcentaje.ToString("0.####", CultureInfo.InvariantCulture)),
                new XAttribute("Activo", x.Activo ? 1 : 0))));

        return xml.ToString(SaveOptions.DisableFormatting);
    }
}
