using System.Data;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AplicacionNotaCreditoRepository(IDbConnectionFactory connectionFactory) : IAplicacionNotaCreditoRepository
{
    public async Task<PagedResult<AplicacionNotaCreditoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<AplicacionNotaCreditoResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_APL_ListarAplicacionesPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Ejercicio", anio);
        command.Parameters.AddWithValue("@Mes", mes);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new AplicacionNotaCreditoResumenDto
            {
                IdAplicacionNotaCredito = reader.GetInt32(reader.GetOrdinal("IdAplicacionNotaCredito")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
                NombrePersona = reader.GetString(reader.GetOrdinal("NombrePersona")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                TipoPersonaTexto = reader.GetString(reader.GetOrdinal("TipoPersonaTexto")),
                FechaAplicacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaAplicacion"))),
                Periodo = reader.GetString(reader.GetOrdinal("Periodo")),
                IdRegistroComprobante = reader.GetInt32(reader.GetOrdinal("IdRegistroComprobante")),
                IdRegistroNotaCredito = reader.GetInt32(reader.GetOrdinal("IdRegistroNotaCredito")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                ImporteAplicado = reader.GetDecimal(reader.GetOrdinal("ImporteAplicado")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                Glosa = reader.GetString(reader.GetOrdinal("Glosa")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? null : reader.GetString(reader.GetOrdinal("Observacion")),
                TipoComprobanteAplicado = reader.GetString(reader.GetOrdinal("TipoComprobanteAplicado")),
                DescripcionTipoComprobanteAplicado = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobanteAplicado")),
                SerieAplicado = reader.GetString(reader.GetOrdinal("SerieAplicado")),
                NumeroAplicado = reader.GetString(reader.GetOrdinal("NumeroAplicado")),
                TipoComprobanteNc = reader.GetString(reader.GetOrdinal("TipoComprobanteNc")),
                DescripcionTipoComprobanteNc = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobanteNc")),
                SerieNc = reader.GetString(reader.GetOrdinal("SerieNc")),
                NumeroNc = reader.GetString(reader.GetOrdinal("NumeroNc"))
            });
        }

        return new PagedResult<AplicacionNotaCreditoResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<IReadOnlyCollection<AplicacionNotaCreditoPendienteDto>> ListarPendientesPorPersonaAsync(int idEmpresa, string moduloOperacion, int idPersona, CancellationToken cancellationToken = default)
    {
        var result = new List<AplicacionNotaCreditoPendienteDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_APL_ListarComprobantesPendientesPorPersona", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@ModuloOperacion", moduloOperacion.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@IdPersona", idPersona);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new AplicacionNotaCreditoPendienteDto
            {
                IdRegistro = reader.GetInt32(reader.GetOrdinal("IdRegistro")),
                ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
                IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
                NombrePersona = reader.GetString(reader.GetOrdinal("NombrePersona")),
                NumeroDocumentoPersona = reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                TipoComprobante = reader.GetString(reader.GetOrdinal("TipoComprobante")),
                DescripcionTipoComprobante = reader.GetString(reader.GetOrdinal("DescripcionTipoComprobante")),
                Serie = reader.GetString(reader.GetOrdinal("Serie")),
                Numero = reader.GetString(reader.GetOrdinal("Numero")),
                IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Saldo = reader.GetDecimal(reader.GetOrdinal("Saldo")),
                EsNotaCredito = reader.GetBoolean(reader.GetOrdinal("EsNotaCredito")),
                EscenarioOperacion = reader.GetString(reader.GetOrdinal("EscenarioOperacion")),
                Observacion = reader.GetString(reader.GetOrdinal("Observacion"))
            });
        }

        return result;
    }

    public async Task<GuardarAplicacionNotaCreditoResultDto> GuardarAsync(GuardarAplicacionNotaCreditoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_APL_GuardarAplicacionNotaCredito", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@ModuloOperacion", request.ModuloOperacion.Trim().ToUpperInvariant());
        command.Parameters.AddWithValue("@IdPersona", request.IdPersona);
        command.Parameters.AddWithValue("@FechaAplicacion", request.FechaAplicacion.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@TipoCambioAplicacion", request.TipoCambio);
        command.Parameters.AddWithValue("@IdRegistroComprobante", request.IdRegistroComprobante);
        command.Parameters.AddWithValue("@IdRegistroNotaCredito", request.IdRegistroNotaCredito);
        command.Parameters.AddWithValue("@ImporteAplicado", request.ImporteAplicado);
        command.Parameters.AddWithValue("@Glosa", request.Glosa.Trim());
        command.Parameters.AddWithValue("@Observacion", string.IsNullOrWhiteSpace(request.Observacion) ? (object)DBNull.Value : request.Observacion.Trim());
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la aplicacion grabada.");
        }

        return new GuardarAplicacionNotaCreditoResultDto
        {
            IdAplicacionNotaCredito = reader.GetInt32(reader.GetOrdinal("IdAplicacionNotaCredito")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            ModuloOperacion = reader.GetString(reader.GetOrdinal("ModuloOperacion")),
            IdPersona = reader.GetInt32(reader.GetOrdinal("IdPersona")),
            FechaAplicacion = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaAplicacion"))),
            IdRegistroComprobante = reader.GetInt32(reader.GetOrdinal("IdRegistroComprobante")),
            IdRegistroNotaCredito = reader.GetInt32(reader.GetOrdinal("IdRegistroNotaCredito")),
            IdMoneda = reader.GetInt32(reader.GetOrdinal("IdMoneda")),
            CodigoMoneda = reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
            ImporteAplicado = reader.GetDecimal(reader.GetOrdinal("ImporteAplicado")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
            Glosa = reader.GetString(reader.GetOrdinal("Glosa"))
        };
    }

    public async Task EliminarAsync(int idAplicacionNotaCredito, int idEmpresa, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_APL_EliminarAplicacionNotaCredito", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdAplicacionNotaCredito", idAplicacionNotaCredito);
        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }
}
