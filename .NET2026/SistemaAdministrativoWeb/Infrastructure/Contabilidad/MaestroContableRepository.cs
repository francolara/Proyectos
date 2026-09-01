using System.Data;
using System.Text.Json;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MaestroContableRepository(IDbConnectionFactory connectionFactory) : IMaestroContableRepository
{
    public async Task<PagedResult<PlanCuentaMaestroDto>> ListarPlanCuentaAsync(string? textoBusqueda, byte? nivelCuenta, bool soloMovimiento, bool soloActivos, int pagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var items = new List<PlanCuentaMaestroDto>();
        var total = 0;
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ListarPlanCuentaMaestro");
        command.Parameters.AddWithValue("@IdPlanCuentaMaestro", DBNull.Value);
        command.Parameters.AddWithValue("@TextoBusqueda", Valor(textoBusqueda));
        command.Parameters.AddWithValue("@NivelCuenta", nivelCuenta.HasValue ? nivelCuenta.Value : DBNull.Value);
        command.Parameters.AddWithValue("@SoloMovimiento", soloMovimiento);
        command.Parameters.AddWithValue("@SoloActivos", soloActivos);
        command.Parameters.AddWithValue("@NumeroPagina", pagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            total = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            items.Add(MapearPlanCuenta(reader));
        }
        return Paginar(items, total, pagina, tamanoPagina);
    }

    public async Task<PlanCuentaMaestroDto?> ObtenerPlanCuentaAsync(int idPlanCuentaMaestro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ListarPlanCuentaMaestro");
        command.Parameters.AddWithValue("@IdPlanCuentaMaestro", idPlanCuentaMaestro);
        command.Parameters.AddWithValue("@TextoBusqueda", DBNull.Value);
        command.Parameters.AddWithValue("@NivelCuenta", DBNull.Value);
        command.Parameters.AddWithValue("@SoloMovimiento", false);
        command.Parameters.AddWithValue("@SoloActivos", false);
        command.Parameters.AddWithValue("@NumeroPagina", 1);
        command.Parameters.AddWithValue("@TamanoPagina", 1);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        return await reader.ReadAsync(cancellationToken) ? MapearPlanCuenta(reader) : null;
    }

    public async Task<int> GuardarPlanCuentaAsync(GuardarPlanCuentaMaestroRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_GuardarPlanCuentaMaestro");
        command.Parameters.AddWithValue("@IdPlanCuentaMaestro", request.IdPlanCuentaMaestro.HasValue ? request.IdPlanCuentaMaestro.Value : DBNull.Value);
        command.Parameters.AddWithValue("@CodigoCuenta", request.CodigoCuenta);
        command.Parameters.AddWithValue("@CodigoCuentaPadre", Valor(request.CodigoCuentaPadre));
        command.Parameters.AddWithValue("@NombreCuenta", request.NombreCuenta);
        command.Parameters.AddWithValue("@ColBalance", request.ColBalance);
        command.Parameters.AddWithValue("@IdMoneda", request.IdMoneda);
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@AceptaMovimiento", request.AceptaMovimiento);
        command.Parameters.AddWithValue("@RequiereCentroCosto", request.RequiereCentroCosto);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@Orden", request.Orden);
        command.Parameters.AddWithValue("@UsuarioRegistro", Valor(request.UsuarioRegistro));
        return await EjecutarIdAsync(connection, command, "IdPlanCuentaMaestro", cancellationToken);
    }

    public Task EliminarPlanCuentaAsync(int idPlanCuentaMaestro, CancellationToken cancellationToken = default) =>
        EjecutarAsync("dbo.usp_CON_EliminarPlanCuentaMaestro", "@IdPlanCuentaMaestro", idPlanCuentaMaestro, cancellationToken);

    public async Task<PagedResult<CuentaDestinoMaestroResumenDto>> ListarCuentasDestinoAsync(string? textoBusqueda, int pagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var items = new List<CuentaDestinoMaestroResumenDto>();
        var total = 0;
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ListarCuentasDestinoMaestro");
        command.Parameters.AddWithValue("@TextoBusqueda", Valor(textoBusqueda));
        command.Parameters.AddWithValue("@NumeroPagina", pagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken))
        {
            total = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            items.Add(new CuentaDestinoMaestroResumenDto
            {
                IdCuentaDestinoReglaMaestro = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoReglaMaestro")),
                CodigoCuentaOrigen = reader.GetString(reader.GetOrdinal("CodigoCuentaOrigen")),
                NombreCuentaOrigen = TextoNullable(reader, "NombreCuentaOrigen"),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
                Observacion = TextoNullable(reader, "Observacion"),
                CantidadTramos = reader.GetInt32(reader.GetOrdinal("CantidadTramos")),
                PorcentajeTotal = reader.GetDecimal(reader.GetOrdinal("PorcentajeTotal"))
            });
        }
        return Paginar(items, total, pagina, tamanoPagina);
    }

    public async Task<CuentaDestinoMaestroDto?> ObtenerCuentaDestinoAsync(int idCuentaDestinoReglaMaestro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ObtenerCuentaDestinoMaestro");
        command.Parameters.AddWithValue("@IdCuentaDestinoReglaMaestro", idCuentaDestinoReglaMaestro);
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) return null;
        var item = new CuentaDestinoMaestroDto
        {
            IdCuentaDestinoReglaMaestro = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoReglaMaestro")),
            CodigoCuentaOrigen = reader.GetString(reader.GetOrdinal("CodigoCuentaOrigen")),
            NombreCuentaOrigen = TextoNullable(reader, "NombreCuentaOrigen"),
            Activo = reader.GetBoolean(reader.GetOrdinal("Activo")),
            Observacion = TextoNullable(reader, "Observacion")
        };
        if (await reader.NextResultAsync(cancellationToken))
        {
            while (await reader.ReadAsync(cancellationToken))
            {
                item.Detalles.Add(new CuentaDestinoDetalleMaestroDto
                {
                    IdCuentaDestinoReglaDetalleMaestro = reader.GetInt32(reader.GetOrdinal("IdCuentaDestinoReglaDetalleMaestro")),
                    Orden = reader.GetInt16(reader.GetOrdinal("Orden")),
                    CodigoCuentaDestinoCargo = reader.GetString(reader.GetOrdinal("CodigoCuentaDestinoCargo")),
                    NombreCuentaDestinoCargo = TextoNullable(reader, "NombreCuentaDestinoCargo"),
                    CodigoCuentaDestinoAbono = reader.GetString(reader.GetOrdinal("CodigoCuentaDestinoAbono")),
                    NombreCuentaDestinoAbono = TextoNullable(reader, "NombreCuentaDestinoAbono"),
                    Porcentaje = reader.GetDecimal(reader.GetOrdinal("Porcentaje")),
                    Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
                });
            }
        }
        return item;
    }

    public async Task<int> GuardarCuentaDestinoAsync(GuardarCuentaDestinoMaestroRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_GuardarCuentaDestinoMaestro");
        command.Parameters.AddWithValue("@IdCuentaDestinoReglaMaestro", request.IdCuentaDestinoReglaMaestro.HasValue ? request.IdCuentaDestinoReglaMaestro.Value : DBNull.Value);
        command.Parameters.AddWithValue("@CodigoCuentaOrigen", request.CodigoCuentaOrigen);
        command.Parameters.AddWithValue("@Activo", request.Activo);
        command.Parameters.AddWithValue("@Observacion", Valor(request.Observacion));
        command.Parameters.AddWithValue("@DetallesJson", JsonSerializer.Serialize(request.Detalles, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase }));
        command.Parameters.AddWithValue("@UsuarioRegistro", Valor(request.UsuarioRegistro));
        return await EjecutarIdAsync(connection, command, "IdCuentaDestinoReglaMaestro", cancellationToken);
    }

    public Task EliminarCuentaDestinoAsync(int idCuentaDestinoReglaMaestro, CancellationToken cancellationToken = default) =>
        EjecutarAsync("dbo.usp_CON_EliminarCuentaDestinoMaestro", "@IdCuentaDestinoReglaMaestro", idCuentaDestinoReglaMaestro, cancellationToken);

    public async Task<AsignacionesMaestroDto> ListarAsignacionesAsync(CancellationToken cancellationToken = default)
    {
        var parametros = new List<ParametroCuentaMaestroDto>();
        var impuestos = new List<TipoImpuestoMaestroDto>();
        var documentos = new List<TipoComprobanteMaestroDto>();
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ListarAsignacionesMaestro");
        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) parametros.Add(new ParametroCuentaMaestroDto
        {
            IdParametroMaestro = reader.GetInt32(reader.GetOrdinal("IdParametroMaestro")), TipoParametro = reader.GetString(reader.GetOrdinal("TipoParametro")),
            CodigoParametro = reader.GetString(reader.GetOrdinal("CodigoParametro")), DescripcionParametro = reader.GetString(reader.GetOrdinal("DescripcionParametro")),
            CodigoCuenta = TextoNullable(reader, "CodigoCuenta"), NombreCuenta = TextoNullable(reader, "NombreCuenta"), Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
        });
        if (await reader.NextResultAsync(cancellationToken)) while (await reader.ReadAsync(cancellationToken)) impuestos.Add(new TipoImpuestoMaestroDto
        {
            IdTipoImpuesto = reader.GetInt32(reader.GetOrdinal("IdTipoImpuesto")), CodigoSunat = reader.GetString(reader.GetOrdinal("CodigoSunat")),
            NombreImpuesto = reader.GetString(reader.GetOrdinal("NombreImpuesto")), CodigoCuenta = TextoNullable(reader, "CodigoCuenta"),
            NombreCuenta = TextoNullable(reader, "NombreCuenta"), Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        });
        if (await reader.NextResultAsync(cancellationToken)) while (await reader.ReadAsync(cancellationToken)) documentos.Add(new TipoComprobanteMaestroDto
        {
            IdTipoComprobante = reader.GetInt32(reader.GetOrdinal("IdTipoComprobante")), CodigoTipoComprobante = reader.GetString(reader.GetOrdinal("CodigoTipoComprobante")),
            Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")), UsoCompras = reader.GetBoolean(reader.GetOrdinal("UsoCompras")), UsoVentas = reader.GetBoolean(reader.GetOrdinal("UsoVentas")),
            CodigoCuentaVentaSoles = TextoNullable(reader, "CodigoCuentaVentaSoles"), NombreCuentaVentaSoles = TextoNullable(reader, "NombreCuentaVentaSoles"),
            CodigoCuentaVentaDolares = TextoNullable(reader, "CodigoCuentaVentaDolares"), NombreCuentaVentaDolares = TextoNullable(reader, "NombreCuentaVentaDolares"),
            CodigoCuentaCompraSoles = TextoNullable(reader, "CodigoCuentaCompraSoles"), NombreCuentaCompraSoles = TextoNullable(reader, "NombreCuentaCompraSoles"),
            CodigoCuentaCompraDolares = TextoNullable(reader, "CodigoCuentaCompraDolares"), NombreCuentaCompraDolares = TextoNullable(reader, "NombreCuentaCompraDolares"),
            Estado = reader.GetBoolean(reader.GetOrdinal("Estado"))
        });
        return new AsignacionesMaestroDto { Parametros = parametros, Impuestos = impuestos, Documentos = documentos };
    }

    public async Task GuardarAsignacionAsync(GuardarAsignacionMaestroRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_GuardarAsignacionMaestro");
        command.Parameters.AddWithValue("@TipoAsignacion", request.TipoAsignacion);
        command.Parameters.AddWithValue("@IdRegistro", request.IdRegistro);
        command.Parameters.AddWithValue("@CodigoCuenta", Valor(request.CodigoCuenta));
        command.Parameters.AddWithValue("@CodigoCuentaVentaSoles", Valor(request.CodigoCuentaVentaSoles));
        command.Parameters.AddWithValue("@CodigoCuentaVentaDolares", Valor(request.CodigoCuentaVentaDolares));
        command.Parameters.AddWithValue("@CodigoCuentaCompraSoles", Valor(request.CodigoCuentaCompraSoles));
        command.Parameters.AddWithValue("@CodigoCuentaCompraDolares", Valor(request.CodigoCuentaCompraDolares));
        command.Parameters.AddWithValue("@UsuarioRegistro", Valor(request.UsuarioRegistro));
        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<PagedResult<OrigenMaestroDto>> ListarOrigenesAsync(string? textoBusqueda, bool soloActivos, int pagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var items = new List<OrigenMaestroDto>(); var total = 0;
        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(connection, "dbo.usp_CON_ListarOrigenesMaestro");
        command.Parameters.AddWithValue("@IdOrigenMaestro", DBNull.Value); command.Parameters.AddWithValue("@TextoBusqueda", Valor(textoBusqueda));
        command.Parameters.AddWithValue("@SoloActivos", soloActivos); command.Parameters.AddWithValue("@NumeroPagina", pagina); command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        await connection.OpenAsync(cancellationToken); await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) { total = reader.GetInt32(reader.GetOrdinal("TotalRegistros")); items.Add(MapearOrigen(reader)); }
        return Paginar(items, total, pagina, tamanoPagina);
    }

    public async Task<OrigenMaestroDto?> ObtenerOrigenAsync(int idOrigenMaestro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, "dbo.usp_CON_ListarOrigenesMaestro");
        command.Parameters.AddWithValue("@IdOrigenMaestro", idOrigenMaestro); command.Parameters.AddWithValue("@TextoBusqueda", DBNull.Value);
        command.Parameters.AddWithValue("@SoloActivos", false); command.Parameters.AddWithValue("@NumeroPagina", 1); command.Parameters.AddWithValue("@TamanoPagina", 1);
        await connection.OpenAsync(cancellationToken); await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        return await reader.ReadAsync(cancellationToken) ? MapearOrigen(reader) : null;
    }

    public async Task<int> GuardarOrigenAsync(GuardarOrigenMaestroRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, "dbo.usp_CON_GuardarOrigenMaestro");
        command.Parameters.AddWithValue("@IdOrigenMaestro", request.IdOrigenMaestro.HasValue ? request.IdOrigenMaestro.Value : DBNull.Value);
        command.Parameters.AddWithValue("@CodigoOrigen", request.CodigoOrigen); command.Parameters.AddWithValue("@NombreOrigen", request.NombreOrigen);
        command.Parameters.AddWithValue("@ModuloOrigen", request.ModuloOrigen); command.Parameters.AddWithValue("@PermiteRegistroManual", request.PermiteRegistroManual);
        command.Parameters.AddWithValue("@Estado", request.Estado); command.Parameters.AddWithValue("@Orden", request.Orden); command.Parameters.AddWithValue("@UsuarioRegistro", Valor(request.UsuarioRegistro));
        return await EjecutarIdAsync(connection, command, "IdOrigenMaestro", cancellationToken);
    }

    public Task EliminarOrigenAsync(int idOrigenMaestro, CancellationToken cancellationToken = default) =>
        EjecutarAsync("dbo.usp_CON_EliminarOrigenMaestro", "@IdOrigenMaestro", idOrigenMaestro, cancellationToken);

    public async Task<PagedResult<ConfiguracionContabilizacionMaestroDto>> ListarConfiguracionesAsync(string? textoBusqueda, int pagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var items = new List<ConfiguracionContabilizacionMaestroDto>(); var total = 0;
        await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, "dbo.usp_CON_ListarConfiguracionContabilizacionMaestro");
        command.Parameters.AddWithValue("@TextoBusqueda", Valor(textoBusqueda)); command.Parameters.AddWithValue("@NumeroPagina", pagina); command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        await connection.OpenAsync(cancellationToken); await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) { total = reader.GetInt32(reader.GetOrdinal("TotalRegistros")); items.Add(new ConfiguracionContabilizacionMaestroDto
        {
            IdConfiguracionContabilizacionMaestro = reader.GetInt32(reader.GetOrdinal("IdConfiguracionContabilizacionMaestro")), ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
            EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")), CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")), NombreOrigen = TextoNullable(reader, "NombreOrigen"),
            Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")), GeneraAsientoAutomatico = reader.GetBoolean(reader.GetOrdinal("GeneraAsientoAutomatico")),
            UsaTipoCambio = reader.GetBoolean(reader.GetOrdinal("UsaTipoCambio")), Activo = reader.GetBoolean(reader.GetOrdinal("Activo")), Orden = reader.GetInt32(reader.GetOrdinal("Orden"))
        }); }
        return Paginar(items, total, pagina, tamanoPagina);
    }

    public async Task GuardarOrigenConfiguracionAsync(int idConfiguracionContabilizacionMaestro, string codigoOrigen, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, "dbo.usp_CON_GuardarOrigenConfiguracionContabilizacionMaestro");
        command.Parameters.AddWithValue("@IdConfiguracionContabilizacionMaestro", idConfiguracionContabilizacionMaestro); command.Parameters.AddWithValue("@CodigoOrigen", codigoOrigen);
        command.Parameters.AddWithValue("@UsuarioRegistro", Valor(usuarioRegistro)); await connection.OpenAsync(cancellationToken); await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<IReadOnlyCollection<ValidacionMaestroIssueDto>> ValidarAsync(CancellationToken cancellationToken = default)
    {
        var items = new List<ValidacionMaestroIssueDto>(); await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, "dbo.usp_CON_ValidarMaestrosContables");
        await connection.OpenAsync(cancellationToken); await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        while (await reader.ReadAsync(cancellationToken)) items.Add(new ValidacionMaestroIssueDto { TipoMaestro = reader.GetString(reader.GetOrdinal("TipoMaestro")), CodigoRegistro = reader.GetString(reader.GetOrdinal("CodigoRegistro")), Descripcion = reader.GetString(reader.GetOrdinal("Descripcion")) });
        return items;
    }

    private async Task EjecutarAsync(string procedimiento, string parametro, int valor, CancellationToken cancellationToken)
    {
        await using var connection = connectionFactory.CreateConnection(); await using var command = CrearComando(connection, procedimiento); command.Parameters.AddWithValue(parametro, valor);
        await connection.OpenAsync(cancellationToken); await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static async Task<int> EjecutarIdAsync(SqlConnection connection, SqlCommand command, string columna, CancellationToken cancellationToken)
    {
        await connection.OpenAsync(cancellationToken); await using var reader = await command.ExecuteReaderAsync(cancellationToken);
        if (!await reader.ReadAsync(cancellationToken)) throw new InvalidOperationException("No se pudo obtener el identificador grabado.");
        return reader.GetInt32(reader.GetOrdinal(columna));
    }

    private static SqlCommand CrearComando(SqlConnection connection, string nombre) => new(nombre, connection) { CommandType = CommandType.StoredProcedure };
    private static object Valor(string? valor) => string.IsNullOrWhiteSpace(valor) ? DBNull.Value : valor.Trim();
    private static string? TextoNullable(SqlDataReader reader, string columna) { var i = reader.GetOrdinal(columna); return reader.IsDBNull(i) ? null : reader.GetString(i); }
    private static PagedResult<T> Paginar<T>(IReadOnlyCollection<T> items, int total, int pagina, int tamanoPagina) => new() { Items = items, TotalRecords = total, PageNumber = pagina, PageSize = tamanoPagina };

    private static PlanCuentaMaestroDto MapearPlanCuenta(SqlDataReader reader) => new()
    {
        IdPlanCuentaMaestro = reader.GetInt32(reader.GetOrdinal("IdPlanCuentaMaestro")), CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
        CodigoCuentaPadre = TextoNullable(reader, "CodigoCuentaPadre"), NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")), NivelCuenta = reader.GetByte(reader.GetOrdinal("NivelCuenta")),
        ColBalance = reader.GetString(reader.GetOrdinal("ColBalance")), IdMoneda = reader.GetString(reader.GetOrdinal("IdMoneda")), TipoCambio = reader.GetString(reader.GetOrdinal("TipoCambio")),
        AceptaMovimiento = reader.GetBoolean(reader.GetOrdinal("AceptaMovimiento")), RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
        Estado = reader.GetBoolean(reader.GetOrdinal("Estado")), Orden = reader.GetInt32(reader.GetOrdinal("Orden")), EsUltimoNivel = reader.GetBoolean(reader.GetOrdinal("EsUltimoNivel"))
    };

    private static OrigenMaestroDto MapearOrigen(SqlDataReader reader) => new()
    {
        IdOrigenMaestro = reader.GetInt32(reader.GetOrdinal("IdOrigenMaestro")), CodigoOrigen = reader.GetString(reader.GetOrdinal("CodigoOrigen")),
        NombreOrigen = reader.GetString(reader.GetOrdinal("NombreOrigen")), ModuloOrigen = reader.GetString(reader.GetOrdinal("ModuloOrigen")),
        PermiteRegistroManual = reader.GetBoolean(reader.GetOrdinal("PermiteRegistroManual")), Estado = reader.GetBoolean(reader.GetOrdinal("Estado")), Orden = reader.GetInt32(reader.GetOrdinal("Orden"))
    };
}
