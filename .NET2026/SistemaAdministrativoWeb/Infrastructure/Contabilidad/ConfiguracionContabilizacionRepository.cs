using System.Data;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContabilizacionRepository(IDbConnectionFactory connectionFactory) : IConfiguracionContabilizacionRepository
{
    public async Task<IReadOnlyCollection<ConfiguracionContabilizacionResumenDto>> ListarPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        var result = new List<ConfiguracionContabilizacionResumenDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarConfiguracionContabilizacionPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new ConfiguracionContabilizacionResumenDto
            {
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                GeneraAsientoAutomatico = reader.GetBoolean(reader.GetOrdinal("GeneraAsientoAutomatico")),
                UsaTipoCambio = reader.GetBoolean(reader.GetOrdinal("UsaTipoCambio")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                CantidadComponentes = reader.GetInt32(reader.GetOrdinal("CantidadComponentes"))
            });
        }

        return result;
    }

    public async Task<PagedResult<ConfiguracionContabilizacionResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<ConfiguracionContabilizacionResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ListarConfiguracionContabilizacionPorEmpresa", connection)
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
            result.Add(new ConfiguracionContabilizacionResumenDto
            {
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                GeneraAsientoAutomatico = reader.GetBoolean(reader.GetOrdinal("GeneraAsientoAutomatico")),
                UsaTipoCambio = reader.GetBoolean(reader.GetOrdinal("UsaTipoCambio")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                CantidadComponentes = reader.GetInt32(reader.GetOrdinal("CantidadComponentes"))
            });
        }

        return new PagedResult<ConfiguracionContabilizacionResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<ConfiguracionContabilizacionDto?> ObtenerAsync(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerConfiguracionContabilizacion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdConfiguracionContabilizacion", idConfiguracionContabilizacion);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        ConfiguracionContabilizacionDto? configuracion = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            configuracion = new ConfiguracionContabilizacionDto
            {
                IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                IdOrigen = reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                GeneraAsientoAutomatico = reader.GetBoolean(reader.GetOrdinal("GeneraAsientoAutomatico")),
                UsaTipoCambio = reader.GetBoolean(reader.GetOrdinal("UsaTipoCambio")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
            };
        }

        if (configuracion is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                configuracion.Detalles.Add(new ConfiguracionContabilizacionDetalleDto
                {
                    IdConfiguracionContabilizacionDetalle = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacionDetalle")),
                    IdConfiguracionContabilizacion = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                    Orden = reader.GetInt16(reader.GetOrdinal("Orden")),
                    ComponenteContable = reader.GetString(reader.GetOrdinal("ComponenteContable")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    NaturalezaMovimiento = reader.GetString(reader.GetOrdinal("NaturalezaMovimiento")),
                    Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
                });
            }
        }

        return configuracion;
    }

    public async Task<int> GuardarAsync(GuardarConfiguracionContabilizacionRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarConfiguracionContabilizacion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdConfiguracionContabilizacion", (object?)request.IdConfiguracionContabilizacion ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@ModuloOperacion", request.ModuloOperacion);
        command.Parameters.AddWithValue("@EscenarioOperacion", request.EscenarioOperacion);
        command.Parameters.AddWithValue("@IdOrigen", request.IdOrigen);
        command.Parameters.AddWithValue("@Descripcion", request.Descripcion);
        command.Parameters.AddWithValue("@GeneraAsientoAutomatico", request.GeneraAsientoAutomatico);
        command.Parameters.AddWithValue("@UsaTipoCambio", request.UsaTipoCambio);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@DetalleXml", ConstruirDetalleXml(request.Detalles));
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)request.UsuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener el identificador de la configuracion grabada.");
        }

        return reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion"));
    }

    public async Task EliminarAsync(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_EliminarConfiguracionContabilizacion", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdConfiguracionContabilizacion", idConfiguracionContabilizacion);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<ConfiguracionContableEmpresaDto> ObtenerConfiguracionContableEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default)
    {
        var result = new ConfiguracionContableEmpresaDto();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_ObtenerConfiguracionContableEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Provisiones.Add(new ConfiguracionContableProvisionDto
            {
                IdConfiguracionContabilizacion = reader.IsDBNull(reader.GetOrdinal("IdConfiguracionContabilizacion")) ? null : reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacion")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                IdOrigen = reader.IsDBNull(reader.GetOrdinal("IdOrigen")) ? null : reader.GetInt32(reader.GetOrdinal("IdOrigen")),
                CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
                NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")),
                GeneraAsientoAutomatico = reader.GetBoolean(reader.GetOrdinal("GeneraAsientoAutomatico")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
            });
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                result.Documentos.Add(new ConfiguracionDocumentoEmpresaDto
                {
                    IdTipoComprobante = reader.GetInt32(reader.GetOrdinal("IdTipoComprobante")),
                    CodigoTipoComprobante = reader.GetString(reader.GetOrdinal("CodigoTipoComprobante")),
                    Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")),
                    UsoCompras = reader.GetBoolean(reader.GetOrdinal("UsoCompras")),
                    UsoVentas = reader.GetBoolean(reader.GetOrdinal("UsoVentas")),
                    IdCuentaVentaSoles = reader.IsDBNull(reader.GetOrdinal("IdCuentaVentaSoles")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaVentaSoles")),
                    CuentaVentaSolesTexto = ConstruirCuentaTexto(reader, "CodigoCuentaVentaSoles", "NombreCuentaVentaSoles"),
                    IdCuentaVentaDolares = reader.IsDBNull(reader.GetOrdinal("IdCuentaVentaDolares")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaVentaDolares")),
                    CuentaVentaDolaresTexto = ConstruirCuentaTexto(reader, "CodigoCuentaVentaDolares", "NombreCuentaVentaDolares"),
                    IdCuentaCompraSoles = reader.IsDBNull(reader.GetOrdinal("IdCuentaCompraSoles")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaCompraSoles")),
                    CuentaCompraSolesTexto = ConstruirCuentaTexto(reader, "CodigoCuentaCompraSoles", "NombreCuentaCompraSoles"),
                    IdCuentaCompraDolares = reader.IsDBNull(reader.GetOrdinal("IdCuentaCompraDolares")) ? null : reader.GetInt32(reader.GetOrdinal("IdCuentaCompraDolares")),
                    CuentaCompraDolaresTexto = ConstruirCuentaTexto(reader, "CodigoCuentaCompraDolares", "NombreCuentaCompraDolares"),
                    Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
                });
            }
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                result.Impuestos.Add(MapearImpuesto(reader));
            }
        }

        return result;
    }

    public async Task GuardarProvisionAsync(int idEmpresa, string moduloOperacion, int idOrigen, bool generaAsientoAutomatico, bool activo, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarProvisionContableEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@ModuloOperacion", moduloOperacion);
        command.Parameters.AddWithValue("@IdOrigen", idOrigen);
        command.Parameters.AddWithValue("@GeneraAsientoAutomatico", generaAsientoAutomatico);
        command.Parameters.AddWithValue("@Activo", activo);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task GuardarDocumentoAsync(int idEmpresa, int idTipoComprobante, int? idCuentaVentaSoles, int? idCuentaVentaDolares, int? idCuentaCompraSoles, int? idCuentaCompraDolares, bool activo, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarDocumentoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdTipoComprobante", idTipoComprobante);
        command.Parameters.AddWithValue("@IdCuentaVentaSoles", (object?)idCuentaVentaSoles ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdCuentaVentaDolares", (object?)idCuentaVentaDolares ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdCuentaCompraSoles", (object?)idCuentaCompraSoles ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdCuentaCompraDolares", (object?)idCuentaCompraDolares ?? DBNull.Value);
        command.Parameters.AddWithValue("@Activo", activo);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task GuardarImpuestoAsync(int idEmpresa, int idTipoImpuesto, int? idPlanCuenta, bool activo, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_GuardarImpuestoConfiguracionEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdTipoImpuesto", idTipoImpuesto);
        command.Parameters.AddWithValue("@IdPlanCuenta", (object?)idPlanCuenta ?? DBNull.Value);
        command.Parameters.AddWithValue("@Activo", activo);
        command.Parameters.AddWithValue("@UsuarioRegistro", (object?)usuarioRegistro ?? DBNull.Value);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static string ConstruirDetalleXml(IReadOnlyCollection<GuardarConfiguracionContabilizacionDetalleRequest> detalles)
    {
        var xml = new XElement("Detalles",
            detalles.Select(x => new XElement("Detalle",
                new XAttribute("Orden", x.Orden),
                new XAttribute("ComponenteContable", x.ComponenteContable),
                new XAttribute("IdPlanCuenta", x.IdPlanCuenta),
                new XAttribute("NaturalezaMovimiento", x.NaturalezaMovimiento),
                new XAttribute("Activo", x.Activo ? 1 : 0))));

        return xml.ToString(SaveOptions.DisableFormatting);
    }

    private static ConfiguracionImpuestoEmpresaDto MapearImpuesto(SqlDataReader reader)
    {
        return new ConfiguracionImpuestoEmpresaDto
        {
            IdTipoImpuesto = reader.GetInt32(reader.GetOrdinal("IdTipoImpuesto")),
            CodigoSunat = reader.GetString(reader.GetOrdinal("CodigoSunat")),
            NombreImpuesto = reader.GetString(reader.GetOrdinal("NombreImpuesto")),
            IdPlanCuenta = reader.IsDBNull(reader.GetOrdinal("IdPlanCuenta")) ? null : reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
            CuentaTexto = ConstruirCuentaTexto(reader, "CodigoCuenta", "NombreCuenta"),
            Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
        };
    }

    private static string ConstruirCuentaTexto(SqlDataReader reader, string codigoColumn, string nombreColumn)
    {
        var codigoOrdinal = reader.GetOrdinal(codigoColumn);
        var nombreOrdinal = reader.GetOrdinal(nombreColumn);

        if (reader.IsDBNull(codigoOrdinal) || reader.IsDBNull(nombreOrdinal))
        {
            return string.Empty;
        }

        return $"{reader.GetString(codigoOrdinal)} - {reader.GetString(nombreOrdinal)}";
    }
}
