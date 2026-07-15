using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroElectronicoRepository(IDbConnectionFactory connectionFactory) : ILibroElectronicoRepository
{
    public Task<IReadOnlyCollection<LibroDiario51Dto>> ListarLibroDiario51Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return EjecutarLibroDiario51Async("dbo.usp_CON_PLE_LibroDiario51_Listar", request, cancellationToken);
    }

    public Task<IReadOnlyCollection<LibroDiario52Dto>> ListarLibroDiario52Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return EjecutarLibroDiario52Async("dbo.usp_CON_PLE_LibroDiario52_Listar", request, cancellationToken);
    }

    public Task<IReadOnlyCollection<LibroMayor61Dto>> ListarLibroMayor61Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return EjecutarLibroMayor61Async("dbo.usp_CON_PLE_LibroMayor61_Listar", request, cancellationToken);
    }

    public async Task<PagedResult<PleHistoryItemDto>> ListarHistorialAsync(int idEmpresa, short anio, byte mes, string? libroElectronico, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<PleHistoryItemDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_PLE_Historial_Listar", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Periodo", PlePeriodoHelper.FormarPeriodoContable(anio, mes));
        command.Parameters.AddWithValue("@LibroElectronico", string.IsNullOrWhiteSpace(libroElectronico) ? (object)DBNull.Value : libroElectronico.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new PleHistoryItemDto
            {
                IdLibroElectronicoGeneracion = reader.GetInt32(reader.GetOrdinal("IdLibroElectronicoGeneracion")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
                CodigoLibro = reader.GetString(reader.GetOrdinal("CodigoLibro")),
                CodigoFormato = reader.GetString(reader.GetOrdinal("CodigoFormato")),
                NombreArchivo = reader.GetString(reader.GetOrdinal("NombreArchivo")),
                CantidadRegistros = reader.GetInt32(reader.GetOrdinal("CantidadRegistros")),
                TotalDebe = reader.GetDecimal(reader.GetOrdinal("TotalDebe")),
                TotalHaber = reader.GetDecimal(reader.GetOrdinal("TotalHaber")),
                Estado = reader.GetString(reader.GetOrdinal("Estado")),
                Observaciones = reader.IsDBNull(reader.GetOrdinal("Observaciones")) ? string.Empty : reader.GetString(reader.GetOrdinal("Observaciones")),
                FechaGeneracion = reader.GetDateTime(reader.GetOrdinal("FechaGeneracion")),
                UsuarioGeneracion = reader.IsDBNull(reader.GetOrdinal("UsuarioGeneracion")) ? string.Empty : reader.GetString(reader.GetOrdinal("UsuarioGeneracion")),
                TotalRegistros = totalRegistros
            });
        }

        return new PagedResult<PleHistoryItemDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task RegistrarHistorialAsync(PleHistorialRegistroRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_CON_PLE_Historial_Registrar", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@Periodo", request.Periodo);
        command.Parameters.AddWithValue("@CodigoLibro", request.CodigoLibro);
        command.Parameters.AddWithValue("@CodigoFormato", request.CodigoFormato);
        command.Parameters.AddWithValue("@NombreArchivo", request.NombreArchivo);
        command.Parameters.AddWithValue("@CantidadRegistros", request.CantidadRegistros);
        command.Parameters.AddWithValue("@TotalDebe", request.TotalDebe);
        command.Parameters.AddWithValue("@TotalHaber", request.TotalHaber);
        command.Parameters.AddWithValue("@Estado", request.Estado);
        command.Parameters.AddWithValue("@Observaciones", string.IsNullOrWhiteSpace(request.Observaciones) ? (object)DBNull.Value : request.Observaciones.Trim());
        command.Parameters.AddWithValue("@UsuarioGeneracion", string.IsNullOrWhiteSpace(request.UsuarioGeneracion) ? (object)DBNull.Value : request.UsuarioGeneracion.Trim());

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private async Task<IReadOnlyCollection<LibroDiario51Dto>> EjecutarLibroDiario51Async(string storedProcedure, LibroElectronicoConsultaRequest request, CancellationToken cancellationToken)
    {
        var result = new List<LibroDiario51Dto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(storedProcedure, request, connection);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new LibroDiario51Dto
            {
                PeriodoPle = reader.GetString(reader.GetOrdinal("PeriodoPle")),
                Cuo = reader.GetString(reader.GetOrdinal("Cuo")),
                CorrelativoAsiento = reader.GetString(reader.GetOrdinal("CorrelativoAsiento")),
                CorrelativoMovimiento = reader.GetString(reader.GetOrdinal("CorrelativoMovimiento")),
                CodigoCuentaContable = reader.GetString(reader.GetOrdinal("CodigoCuentaContable")),
                CodigoUnidadOperacion = reader.IsDBNull(reader.GetOrdinal("CodigoUnidadOperacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoUnidadOperacion")),
                CodigoCentroCosto = reader.IsDBNull(reader.GetOrdinal("CodigoCentroCosto")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoCentroCosto")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                CodigoLibroRelacionado = reader.IsDBNull(reader.GetOrdinal("CodigoLibroRelacionado")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoLibroRelacionado")),
                TipoDocumentoEmisor = reader.IsDBNull(reader.GetOrdinal("TipoDocumentoEmisor")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoDocumentoEmisor")),
                NumeroDocumentoEmisor = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoEmisor")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoEmisor")),
                TipoComprobante = reader.IsDBNull(reader.GetOrdinal("TipoComprobante")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoComprobante")),
                SerieComprobante = reader.IsDBNull(reader.GetOrdinal("SerieComprobante")) ? string.Empty : reader.GetString(reader.GetOrdinal("SerieComprobante")),
                NumeroComprobante = reader.IsDBNull(reader.GetOrdinal("NumeroComprobante")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroComprobante")),
                FechaContable = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaContable"))),
                FechaVencimiento = reader.IsDBNull(reader.GetOrdinal("FechaVencimiento")) ? null : DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaVencimiento"))),
                FechaOperacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaOperacion"))),
                Glosa = reader.IsDBNull(reader.GetOrdinal("Glosa")) ? string.Empty : reader.GetString(reader.GetOrdinal("Glosa")),
                GlosaReferencial = reader.IsDBNull(reader.GetOrdinal("GlosaReferencial")) ? string.Empty : reader.GetString(reader.GetOrdinal("GlosaReferencial")),
                Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                InformacionComplementaria = reader.IsDBNull(reader.GetOrdinal("InformacionComplementaria")) ? string.Empty : reader.GetString(reader.GetOrdinal("InformacionComplementaria")),
                EstadoOperacion = reader.GetString(reader.GetOrdinal("EstadoOperacion")),
                NumeroAsiento = reader.GetInt32(reader.GetOrdinal("NumeroAsiento"))
            });
        }

        return result;
    }

    private async Task<IReadOnlyCollection<LibroDiario52Dto>> EjecutarLibroDiario52Async(string storedProcedure, LibroElectronicoConsultaRequest request, CancellationToken cancellationToken)
    {
        var result = new List<LibroDiario52Dto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(storedProcedure, request, connection);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new LibroDiario52Dto
            {
                PeriodoPle = reader.GetString(reader.GetOrdinal("PeriodoPle")),
                Cuo = reader.GetString(reader.GetOrdinal("Cuo")),
                CorrelativoAsiento = reader.GetString(reader.GetOrdinal("CorrelativoAsiento")),
                FechaOperacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaOperacion"))),
                Glosa = reader.IsDBNull(reader.GetOrdinal("Glosa")) ? string.Empty : reader.GetString(reader.GetOrdinal("Glosa")),
                CodigoCuentaContable = reader.GetString(reader.GetOrdinal("CodigoCuentaContable")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                EstadoOperacion = reader.GetString(reader.GetOrdinal("EstadoOperacion")),
                NumeroAsiento = reader.GetInt32(reader.GetOrdinal("NumeroAsiento"))
            });
        }

        return result;
    }

    private async Task<IReadOnlyCollection<LibroMayor61Dto>> EjecutarLibroMayor61Async(string storedProcedure, LibroElectronicoConsultaRequest request, CancellationToken cancellationToken)
    {
        var result = new List<LibroMayor61Dto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = CrearComando(storedProcedure, request, connection);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new LibroMayor61Dto
            {
                PeriodoPle = reader.GetString(reader.GetOrdinal("PeriodoPle")),
                Cuo = reader.GetString(reader.GetOrdinal("Cuo")),
                CorrelativoMovimiento = reader.GetString(reader.GetOrdinal("CorrelativoMovimiento")),
                CodigoCuentaContable = reader.GetString(reader.GetOrdinal("CodigoCuentaContable")),
                FechaOperacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaOperacion"))),
                Glosa = reader.IsDBNull(reader.GetOrdinal("Glosa")) ? string.Empty : reader.GetString(reader.GetOrdinal("Glosa")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                EstadoOperacion = reader.GetString(reader.GetOrdinal("EstadoOperacion")),
                NumeroAsiento = reader.GetInt32(reader.GetOrdinal("NumeroAsiento"))
            });
        }

        return result;
    }

    private static SqlCommand CrearComando(string storedProcedure, LibroElectronicoConsultaRequest request, SqlConnection connection)
    {
        var command = new SqlCommand(storedProcedure, connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdAnno", request.Anio);
        command.Parameters.AddWithValue("@Mes", request.Mes);
        command.Parameters.AddWithValue("@Moneda", "PEN");
        command.Parameters.AddWithValue("@Estado", string.IsNullOrWhiteSpace(request.Estado) ? (object)DBNull.Value : request.Estado.Trim());
        command.Parameters.AddWithValue("@FechaDesde", request.FechaDesde.HasValue ? request.FechaDesde.Value.ToDateTime(TimeOnly.MinValue) : (object)DBNull.Value);
        command.Parameters.AddWithValue("@FechaHasta", request.FechaHasta.HasValue ? request.FechaHasta.Value.ToDateTime(TimeOnly.MinValue) : (object)DBNull.Value);

        return command;
    }
}
