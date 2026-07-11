using System.Data;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;
using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CajaBancoRepository(IDbConnectionFactory connectionFactory) : ICajaBancoRepository
{
    public async Task<PagedResult<CajaBancoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, int? idBancoConfiguracionEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<CajaBancoResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_ListarMovimientosBancoPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresa", (object?)idBancoConfiguracionEmpresa ?? DBNull.Value);
        command.Parameters.AddWithValue("@Anio", anio);
        command.Parameters.AddWithValue("@Mes", mes);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(MapResumen(reader));
        }

        return new PagedResult<CajaBancoResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<CajaBancoResumenCuentaDto> ObtenerResumenCuentaAsync(int idEmpresa, int? idBancoConfiguracionEmpresa, short anio, byte mes, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_ObtenerResumenMovimientoBanco", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresa", (object?)idBancoConfiguracionEmpresa ?? DBNull.Value);
        command.Parameters.AddWithValue("@Anio", anio);
        command.Parameters.AddWithValue("@Mes", mes);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            return new CajaBancoResumenCuentaDto();
        }

        return new CajaBancoResumenCuentaDto
        {
            SaldoInicial = reader.GetDecimal(reader.GetOrdinal("SaldoInicial")),
            IngresosMes = reader.GetDecimal(reader.GetOrdinal("IngresosMes")),
            EgresosMes = reader.GetDecimal(reader.GetOrdinal("EgresosMes")),
            SaldoFinal = reader.GetDecimal(reader.GetOrdinal("SaldoFinal"))
        };
    }

    public async Task<IReadOnlyCollection<OperacionBancariaDto>> ListarOperacionesBancariasAsync(string tipoMovimiento, string? textoBusqueda = null, int tamanoPagina = 100, string? idTipoOpeBancaria = null, CancellationToken cancellationToken = default)
    {
        var result = new List<OperacionBancariaDto>();

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_ListarOperacionesBancarias", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@TipoMovimiento", tipoMovimiento);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);
        command.Parameters.AddWithValue("@IdTipoOpeBancaria", string.IsNullOrWhiteSpace(idTipoOpeBancaria) ? (object)DBNull.Value : idTipoOpeBancaria.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            result.Add(new OperacionBancariaDto
            {
                IdOpeBancaria = reader.GetString(reader.GetOrdinal("IdOpeBancaria")),
                IdTipoOpeBancaria = reader.IsDBNull(reader.GetOrdinal("IdTipoOpeBancaria")) ? string.Empty : reader.GetString(reader.GetOrdinal("IdTipoOpeBancaria")),
                TipoMovimiento = reader.GetString(reader.GetOrdinal("TipoMovimiento")),
                TipoOperacion = reader.GetString(reader.GetOrdinal("TipoOperacion"))
            });
        }

        return result;
    }

    public async Task<CajaBancoMovimientoDto?> ObtenerPorIdAsync(int idEmpresa, int idMovimientoBanco, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_ObtenerMovimientoBanco", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdMovimientoBanco", idMovimientoBanco);
        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        CajaBancoMovimientoDto? movimiento = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            movimiento = new CajaBancoMovimientoDto
            {
                IdMovimientoBanco = reader.GetInt32(reader.GetOrdinal("IdMovimientoBanco")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                NumeroMovimiento = reader.GetInt32(reader.GetOrdinal("NumeroMovimiento")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdBancoConfiguracionEmpresa = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresa")),
                NroCuentaCorriente = reader.GetString(reader.GetOrdinal("NroCuentaCorriente")),
                CodigoBanco = reader.GetString(reader.GetOrdinal("CodigoBanco")),
                NombreBanco = reader.GetString(reader.GetOrdinal("NombreBanco")),
                TitularCuentaCorriente = reader.IsDBNull(reader.GetOrdinal("TitularCuentaCorriente")) ? string.Empty : reader.GetString(reader.GetOrdinal("TitularCuentaCorriente")),
                IdMoneda = reader.IsDBNull(reader.GetOrdinal("IdMoneda")) ? null : reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.IsDBNull(reader.GetOrdinal("CodigoMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                NombreMoneda = reader.IsDBNull(reader.GetOrdinal("NombreMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreMoneda")),
                TipoMovimiento = reader.GetString(reader.GetOrdinal("TipoMovimiento")),
                IdOpeBancaria = reader.GetString(reader.GetOrdinal("IdOpeBancaria")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                IdPersona = reader.IsDBNull(reader.GetOrdinal("IdPersona")) ? null : reader.GetInt32(reader.GetOrdinal("IdPersona")),
                NumeroDocumentoPersona = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoPersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                NombrePersona = reader.IsDBNull(reader.GetOrdinal("NombrePersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombrePersona")),
                NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                Glosa = reader.GetString(reader.GetOrdinal("Glosa")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("Observacion")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
            };
        }

        if (movimiento is null)
        {
            return null;
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            var detalles = new List<CajaBancoDetalleDto>();
            while (await reader.ReadAsync(cancellationToken))
            {
                detalles.Add(new CajaBancoDetalleDto
                {
                    IdMovimientoBancoDetalle = reader.IsDBNull(reader.GetOrdinal("IdMovimientoBancoDetalle")) ? null : reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoDetalle")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    IdPersona = reader.IsDBNull(reader.GetOrdinal("IdPersona")) ? null : reader.GetInt32(reader.GetOrdinal("IdPersona")),
                    ModuloOperacionComprobante = reader.IsDBNull(reader.GetOrdinal("ModuloOperacionComprobante")) ? null : reader.GetString(reader.GetOrdinal("ModuloOperacionComprobante")),
                    IdRegistroComprobante = reader.IsDBNull(reader.GetOrdinal("IdRegistroComprobante")) ? null : reader.GetInt32(reader.GetOrdinal("IdRegistroComprobante")),
                    ImporteAplicado = reader.IsDBNull(reader.GetOrdinal("ImporteAplicado")) ? null : reader.GetDecimal(reader.GetOrdinal("ImporteAplicado")),
                    NumeroDocumentoPersona = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoPersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                    NombrePersona = reader.IsDBNull(reader.GetOrdinal("NombrePersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombrePersona")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
                    GlosaDetalle = reader.IsDBNull(reader.GetOrdinal("GlosaDetalle")) ? null : reader.GetString(reader.GetOrdinal("GlosaDetalle")),
                    CodigoCentroCosto = reader.IsDBNull(reader.GetOrdinal("CodigoCentroCosto")) ? null : reader.GetString(reader.GetOrdinal("CodigoCentroCosto")),
                    NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? null : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                    TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? null : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                    Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? null : reader.GetString(reader.GetOrdinal("Serie")),
                    ReferenciaLinea = reader.IsDBNull(reader.GetOrdinal("ReferenciaLinea")) ? null : reader.GetString(reader.GetOrdinal("ReferenciaLinea")),
                    TipoCambioLinea = reader.IsDBNull(reader.GetOrdinal("TipoCambioLinea")) ? null : reader.GetDecimal(reader.GetOrdinal("TipoCambioLinea")),
                    Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                    Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                    TotalImporteS = reader.IsDBNull(reader.GetOrdinal("TotalImporteS")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TotalImporteS")),
                    TotalImporteD = reader.IsDBNull(reader.GetOrdinal("TotalImporteD")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TotalImporteD"))
                });
            }

            movimiento = new CajaBancoMovimientoDto
            {
                IdMovimientoBanco = movimiento.IdMovimientoBanco,
                IdAsiento = movimiento.IdAsiento,
                NumeroAsiento = movimiento.NumeroAsiento,
                NumeroMovimiento = movimiento.NumeroMovimiento,
                IdEmpresa = movimiento.IdEmpresa,
                IdBancoConfiguracionEmpresa = movimiento.IdBancoConfiguracionEmpresa,
                NroCuentaCorriente = movimiento.NroCuentaCorriente,
                CodigoBanco = movimiento.CodigoBanco,
                NombreBanco = movimiento.NombreBanco,
                TitularCuentaCorriente = movimiento.TitularCuentaCorriente,
                IdMoneda = movimiento.IdMoneda,
                CodigoMoneda = movimiento.CodigoMoneda,
                NombreMoneda = movimiento.NombreMoneda,
                TipoMovimiento = movimiento.TipoMovimiento,
                IdOpeBancaria = movimiento.IdOpeBancaria,
                FechaEmision = movimiento.FechaEmision,
                TipoCambio = movimiento.TipoCambio,
                IdPersona = movimiento.IdPersona,
                NumeroDocumentoPersona = movimiento.NumeroDocumentoPersona,
                NombrePersona = movimiento.NombrePersona,
                NumeroDocumento = movimiento.NumeroDocumento,
                Glosa = movimiento.Glosa,
                Observacion = movimiento.Observacion,
                ImporteTotal = movimiento.ImporteTotal,
                Activo = movimiento.Activo,
                Detalles = detalles
            };
        }

        return movimiento;
    }

    public async Task<CajaBancoMovimientoDto> GuardarAsync(GuardarCajaBancoRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_GuardarMovimientoBanco", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        var detallesXml = new XDocument(
            new XElement("Detalles",
                request.Detalles.Select(x => new XElement("Detalle",
                    new XAttribute("Item", x.Item),
                    new XAttribute("IdPlanCuenta", x.IdPlanCuenta),
                    new XAttribute("IdPersona", x.IdPersona ?? 0),
                    new XAttribute("ModuloOperacionComprobante", x.ModuloOperacionComprobante ?? string.Empty),
                    new XAttribute("IdRegistroComprobante", x.IdRegistroComprobante ?? 0),
                    new XAttribute("ImporteAplicado", x.ImporteAplicado ?? 0),
                    new XAttribute("GlosaDetalle", x.GlosaDetalle ?? string.Empty),
                    new XAttribute("CodigoCentroCosto", x.CodigoCentroCosto ?? string.Empty),
                    new XAttribute("NumeroDocumento", x.NumeroDocumento ?? string.Empty),
                    new XAttribute("TipoDocumento", x.TipoDocumento ?? string.Empty),
                    new XAttribute("Serie", x.Serie ?? string.Empty),
                    new XAttribute("ReferenciaLinea", x.ReferenciaLinea ?? string.Empty),
                    new XAttribute("TipoCambioLinea", x.TipoCambioLinea),
                    new XAttribute("Debe", x.Debe),
                    new XAttribute("Haber", x.Haber)))));

        command.Parameters.AddWithValue("@IdMovimientoBanco", (object?)request.IdMovimientoBanco ?? DBNull.Value);
        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresa", request.IdBancoConfiguracionEmpresa);
        command.Parameters.AddWithValue("@TipoMovimiento", request.TipoMovimiento);
        command.Parameters.AddWithValue("@IdOpeBancaria", request.IdOpeBancaria);
        command.Parameters.AddWithValue("@FechaEmision", request.FechaEmision.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@TipoCambio", request.TipoCambio);
        command.Parameters.AddWithValue("@IdPersona", (object?)request.IdPersona ?? DBNull.Value);
        command.Parameters.AddWithValue("@NumeroDocumento", string.IsNullOrWhiteSpace(request.NumeroDocumento) ? (object)DBNull.Value : request.NumeroDocumento.Trim());
        command.Parameters.AddWithValue("@Glosa", request.Glosa.Trim());
        command.Parameters.AddWithValue("@Observacion", string.IsNullOrWhiteSpace(request.Observacion) ? (object)DBNull.Value : request.Observacion.Trim());
        command.Parameters.AddWithValue("@ImporteTotal", request.ImporteTotal);
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());
        command.Parameters.AddWithValue("@DetallesXml", detallesXml.ToString(SaveOptions.DisableFormatting));

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        CajaBancoMovimientoDto? movimiento = null;

        if (await reader.ReadAsync(cancellationToken))
        {
            movimiento = new CajaBancoMovimientoDto
            {
                IdMovimientoBanco = reader.GetInt32(reader.GetOrdinal("IdMovimientoBanco")),
                IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
                NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
                NumeroMovimiento = reader.GetInt32(reader.GetOrdinal("NumeroMovimiento")),
                IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
                IdBancoConfiguracionEmpresa = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresa")),
                NroCuentaCorriente = reader.GetString(reader.GetOrdinal("NroCuentaCorriente")),
                CodigoBanco = reader.GetString(reader.GetOrdinal("CodigoBanco")),
                NombreBanco = reader.GetString(reader.GetOrdinal("NombreBanco")),
                TitularCuentaCorriente = reader.IsDBNull(reader.GetOrdinal("TitularCuentaCorriente")) ? string.Empty : reader.GetString(reader.GetOrdinal("TitularCuentaCorriente")),
                IdMoneda = reader.IsDBNull(reader.GetOrdinal("IdMoneda")) ? null : reader.GetInt32(reader.GetOrdinal("IdMoneda")),
                CodigoMoneda = reader.IsDBNull(reader.GetOrdinal("CodigoMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoMoneda")),
                NombreMoneda = reader.IsDBNull(reader.GetOrdinal("NombreMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreMoneda")),
                TipoMovimiento = reader.GetString(reader.GetOrdinal("TipoMovimiento")),
                IdOpeBancaria = reader.GetString(reader.GetOrdinal("IdOpeBancaria")),
                FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
                TipoCambio = reader.GetDecimal(reader.GetOrdinal("TipoCambio")),
                IdPersona = reader.IsDBNull(reader.GetOrdinal("IdPersona")) ? null : reader.GetInt32(reader.GetOrdinal("IdPersona")),
                NumeroDocumentoPersona = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoPersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                NombrePersona = reader.IsDBNull(reader.GetOrdinal("NombrePersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombrePersona")),
                NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                Glosa = reader.GetString(reader.GetOrdinal("Glosa")),
                Observacion = reader.IsDBNull(reader.GetOrdinal("Observacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("Observacion")),
                ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
                Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
            };
        }

        if (movimiento is null)
        {
            throw new InvalidOperationException("No se pudo obtener el movimiento grabado.");
        }

        if (await reader.NextResultAsync(cancellationToken))
        {
            var detalles = new List<CajaBancoDetalleDto>();
            while (await reader.ReadAsync(cancellationToken))
            {
                detalles.Add(new CajaBancoDetalleDto
                {
                    IdMovimientoBancoDetalle = reader.IsDBNull(reader.GetOrdinal("IdMovimientoBancoDetalle")) ? null : reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoDetalle")),
                    Item = reader.GetInt16(reader.GetOrdinal("Item")),
                    IdPlanCuenta = reader.GetInt32(reader.GetOrdinal("IdPlanCuenta")),
                    IdPersona = reader.IsDBNull(reader.GetOrdinal("IdPersona")) ? null : reader.GetInt32(reader.GetOrdinal("IdPersona")),
                    ModuloOperacionComprobante = reader.IsDBNull(reader.GetOrdinal("ModuloOperacionComprobante")) ? null : reader.GetString(reader.GetOrdinal("ModuloOperacionComprobante")),
                    IdRegistroComprobante = reader.IsDBNull(reader.GetOrdinal("IdRegistroComprobante")) ? null : reader.GetInt32(reader.GetOrdinal("IdRegistroComprobante")),
                    ImporteAplicado = reader.IsDBNull(reader.GetOrdinal("ImporteAplicado")) ? null : reader.GetDecimal(reader.GetOrdinal("ImporteAplicado")),
                    NumeroDocumentoPersona = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoPersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
                    NombrePersona = reader.IsDBNull(reader.GetOrdinal("NombrePersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombrePersona")),
                    CodigoCuenta = reader.GetString(reader.GetOrdinal("CodigoCuenta")),
                    NombreCuenta = reader.GetString(reader.GetOrdinal("NombreCuenta")),
                    RequiereCentroCosto = reader.GetBoolean(reader.GetOrdinal("RequiereCentroCosto")),
                    GlosaDetalle = reader.IsDBNull(reader.GetOrdinal("GlosaDetalle")) ? null : reader.GetString(reader.GetOrdinal("GlosaDetalle")),
                    CodigoCentroCosto = reader.IsDBNull(reader.GetOrdinal("CodigoCentroCosto")) ? null : reader.GetString(reader.GetOrdinal("CodigoCentroCosto")),
                    NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? null : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
                    TipoDocumento = reader.IsDBNull(reader.GetOrdinal("TipoDocumento")) ? null : reader.GetString(reader.GetOrdinal("TipoDocumento")),
                    Serie = reader.IsDBNull(reader.GetOrdinal("Serie")) ? null : reader.GetString(reader.GetOrdinal("Serie")),
                    ReferenciaLinea = reader.IsDBNull(reader.GetOrdinal("ReferenciaLinea")) ? null : reader.GetString(reader.GetOrdinal("ReferenciaLinea")),
                    TipoCambioLinea = reader.IsDBNull(reader.GetOrdinal("TipoCambioLinea")) ? null : reader.GetDecimal(reader.GetOrdinal("TipoCambioLinea")),
                    Debe = reader.GetDecimal(reader.GetOrdinal("Debe")),
                    Haber = reader.GetDecimal(reader.GetOrdinal("Haber")),
                    TotalImporteS = reader.IsDBNull(reader.GetOrdinal("TotalImporteS")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TotalImporteS")),
                    TotalImporteD = reader.IsDBNull(reader.GetOrdinal("TotalImporteD")) ? 0m : reader.GetDecimal(reader.GetOrdinal("TotalImporteD"))
                });
            }

            movimiento = new CajaBancoMovimientoDto
            {
                IdMovimientoBanco = movimiento.IdMovimientoBanco,
                IdAsiento = movimiento.IdAsiento,
                NumeroAsiento = movimiento.NumeroAsiento,
                NumeroMovimiento = movimiento.NumeroMovimiento,
                IdEmpresa = movimiento.IdEmpresa,
                IdBancoConfiguracionEmpresa = movimiento.IdBancoConfiguracionEmpresa,
                NroCuentaCorriente = movimiento.NroCuentaCorriente,
                CodigoBanco = movimiento.CodigoBanco,
                NombreBanco = movimiento.NombreBanco,
                TitularCuentaCorriente = movimiento.TitularCuentaCorriente,
                IdMoneda = movimiento.IdMoneda,
                CodigoMoneda = movimiento.CodigoMoneda,
                NombreMoneda = movimiento.NombreMoneda,
                TipoMovimiento = movimiento.TipoMovimiento,
                IdOpeBancaria = movimiento.IdOpeBancaria,
                FechaEmision = movimiento.FechaEmision,
                TipoCambio = movimiento.TipoCambio,
                IdPersona = movimiento.IdPersona,
                NumeroDocumentoPersona = movimiento.NumeroDocumentoPersona,
                NombrePersona = movimiento.NombrePersona,
                NumeroDocumento = movimiento.NumeroDocumento,
                Glosa = movimiento.Glosa,
                Observacion = movimiento.Observacion,
                ImporteTotal = movimiento.ImporteTotal,
                Activo = movimiento.Activo,
                Detalles = detalles
            };
        }

        return movimiento;
    }

    public async Task EliminarAsync(int idEmpresa, int idMovimientoBanco, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_EliminarMovimientoBanco", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdMovimientoBanco", idMovimientoBanco);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    public async Task<PagedResult<TransferenciaCuentaResumenDto>> ListarTransferenciasPaginadoPorEmpresaAsync(int idEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default)
    {
        var result = new List<TransferenciaCuentaResumenDto>();
        var totalRegistros = 0;

        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_ListarTransferenciasCuentaPorEmpresa", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@Anio", anio);
        command.Parameters.AddWithValue("@Mes", mes);
        command.Parameters.AddWithValue("@TextoBusqueda", string.IsNullOrWhiteSpace(textoBusqueda) ? (object)DBNull.Value : textoBusqueda.Trim());
        command.Parameters.AddWithValue("@NumeroPagina", numeroPagina);
        command.Parameters.AddWithValue("@TamanoPagina", tamanoPagina);

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        while (await reader.ReadAsync(cancellationToken))
        {
            totalRegistros = reader.GetInt32(reader.GetOrdinal("TotalRegistros"));
            result.Add(new TransferenciaCuentaResumenDto
            {
                IdTransferenciaCuenta = reader.GetGuid(reader.GetOrdinal("IdTransferenciaCuenta")),
                IdMovimientoBancoEmisor = reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoEmisor")),
                IdAsientoEmisor = reader.IsDBNull(reader.GetOrdinal("IdAsientoEmisor")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsientoEmisor")),
                NumeroMovimientoEmisor = reader.GetInt32(reader.GetOrdinal("NumeroMovimientoEmisor")),
                NumeroAsientoEmisor = reader.IsDBNull(reader.GetOrdinal("NumeroAsientoEmisor")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsientoEmisor")),
                IdBancoConfiguracionEmpresaEmisor = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresaEmisor")),
                CuentaCorrienteEmisor = reader.GetString(reader.GetOrdinal("CuentaCorrienteEmisor")),
                MonedaEmisor = reader.IsDBNull(reader.GetOrdinal("MonedaEmisor")) ? string.Empty : reader.GetString(reader.GetOrdinal("MonedaEmisor")),
                OperacionEmisor = reader.IsDBNull(reader.GetOrdinal("OperacionEmisor")) ? string.Empty : reader.GetString(reader.GetOrdinal("OperacionEmisor")),
                FechaEmisionEmisor = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmisionEmisor"))),
                TipoCambioEmisor = reader.GetDecimal(reader.GetOrdinal("TipoCambioEmisor")),
                NumeroOperacionEmisor = reader.IsDBNull(reader.GetOrdinal("NumeroOperacionEmisor")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroOperacionEmisor")),
                ImporteEmisor = reader.GetDecimal(reader.GetOrdinal("ImporteEmisor")),
                GlosaEmisor = reader.GetString(reader.GetOrdinal("GlosaEmisor")),
                IdMovimientoBancoReceptor = reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoReceptor")),
                IdAsientoReceptor = reader.IsDBNull(reader.GetOrdinal("IdAsientoReceptor")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsientoReceptor")),
                NumeroMovimientoReceptor = reader.GetInt32(reader.GetOrdinal("NumeroMovimientoReceptor")),
                NumeroAsientoReceptor = reader.IsDBNull(reader.GetOrdinal("NumeroAsientoReceptor")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsientoReceptor")),
                IdBancoConfiguracionEmpresaReceptor = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresaReceptor")),
                CuentaCorrienteReceptor = reader.GetString(reader.GetOrdinal("CuentaCorrienteReceptor")),
                MonedaReceptor = reader.IsDBNull(reader.GetOrdinal("MonedaReceptor")) ? string.Empty : reader.GetString(reader.GetOrdinal("MonedaReceptor")),
                OperacionReceptor = reader.IsDBNull(reader.GetOrdinal("OperacionReceptor")) ? string.Empty : reader.GetString(reader.GetOrdinal("OperacionReceptor")),
                FechaEmisionReceptor = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmisionReceptor"))),
                TipoCambioReceptor = reader.GetDecimal(reader.GetOrdinal("TipoCambioReceptor")),
                NumeroOperacionReceptor = reader.IsDBNull(reader.GetOrdinal("NumeroOperacionReceptor")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroOperacionReceptor")),
                ImporteReceptor = reader.GetDecimal(reader.GetOrdinal("ImporteReceptor")),
                GlosaReceptor = reader.GetString(reader.GetOrdinal("GlosaReceptor"))
            });
        }

        return new PagedResult<TransferenciaCuentaResumenDto>
        {
            Items = result,
            TotalRecords = totalRegistros,
            PageNumber = numeroPagina,
            PageSize = tamanoPagina
        };
    }

    public async Task<TransferenciaCuentaResultadoDto> GuardarTransferenciaAsync(GuardarTransferenciaCuentaRequest request, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_GuardarTransferenciaCuenta", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", request.IdEmpresa);
        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresaEmisor", request.IdBancoConfiguracionEmpresaEmisor);
        command.Parameters.AddWithValue("@IdBancoConfiguracionEmpresaReceptor", request.IdBancoConfiguracionEmpresaReceptor);
        command.Parameters.AddWithValue("@IdOpeBancariaEmisor", request.IdOpeBancariaEmisor);
        command.Parameters.AddWithValue("@IdOpeBancariaReceptor", request.IdOpeBancariaReceptor);
        command.Parameters.AddWithValue("@FechaEmisionEmisor", request.FechaEmisionEmisor.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@FechaEmisionReceptor", request.FechaEmisionReceptor.ToDateTime(TimeOnly.MinValue));
        command.Parameters.AddWithValue("@TipoCambioEmisor", request.TipoCambioEmisor);
        command.Parameters.AddWithValue("@TipoCambioReceptor", request.TipoCambioReceptor);
        command.Parameters.AddWithValue("@NumeroOperacionEmisor", string.IsNullOrWhiteSpace(request.NumeroOperacionEmisor) ? (object)DBNull.Value : request.NumeroOperacionEmisor.Trim());
        command.Parameters.AddWithValue("@NumeroOperacionReceptor", string.IsNullOrWhiteSpace(request.NumeroOperacionReceptor) ? (object)DBNull.Value : request.NumeroOperacionReceptor.Trim());
        command.Parameters.AddWithValue("@ImporteEmisor", request.ImporteEmisor);
        command.Parameters.AddWithValue("@ImporteReceptor", request.ImporteReceptor);
        command.Parameters.AddWithValue("@GlosaEmisor", request.GlosaEmisor.Trim());
        command.Parameters.AddWithValue("@GlosaReceptor", request.GlosaReceptor.Trim());
        command.Parameters.AddWithValue("@ObservacionEmisor", string.IsNullOrWhiteSpace(request.ObservacionEmisor) ? (object)DBNull.Value : request.ObservacionEmisor.Trim());
        command.Parameters.AddWithValue("@ObservacionReceptor", string.IsNullOrWhiteSpace(request.ObservacionReceptor) ? (object)DBNull.Value : request.ObservacionReceptor.Trim());
        command.Parameters.AddWithValue("@UsuarioRegistro", string.IsNullOrWhiteSpace(request.UsuarioRegistro) ? (object)DBNull.Value : request.UsuarioRegistro.Trim());

        await connection.OpenAsync(cancellationToken);
        await using var reader = await command.ExecuteReaderAsync(cancellationToken);

        if (!await reader.ReadAsync(cancellationToken))
        {
            throw new InvalidOperationException("No se pudo obtener la transferencia grabada.");
        }

        return new TransferenciaCuentaResultadoDto
        {
            IdTransferenciaCuenta = reader.GetGuid(reader.GetOrdinal("IdTransferenciaCuenta")),
            IdMovimientoBancoEmisor = reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoEmisor")),
            NumeroMovimientoEmisor = reader.GetInt32(reader.GetOrdinal("NumeroMovimientoEmisor")),
            NumeroAsientoEmisor = reader.IsDBNull(reader.GetOrdinal("NumeroAsientoEmisor")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsientoEmisor")),
            IdMovimientoBancoReceptor = reader.GetInt32(reader.GetOrdinal("IdMovimientoBancoReceptor")),
            NumeroMovimientoReceptor = reader.GetInt32(reader.GetOrdinal("NumeroMovimientoReceptor")),
            NumeroAsientoReceptor = reader.IsDBNull(reader.GetOrdinal("NumeroAsientoReceptor")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsientoReceptor")),
            ImporteEmisor = reader.GetDecimal(reader.GetOrdinal("ImporteEmisor")),
            ImporteReceptor = reader.GetDecimal(reader.GetOrdinal("ImporteReceptor"))
        };
    }

    public async Task EliminarTransferenciaAsync(int idEmpresa, int idMovimientoBancoEmisor, CancellationToken cancellationToken = default)
    {
        await using var connection = connectionFactory.CreateConnection();
        await using var command = new SqlCommand("dbo.usp_BAN_EliminarTransferenciaCuenta", connection)
        {
            CommandType = CommandType.StoredProcedure
        };

        command.Parameters.AddWithValue("@IdEmpresa", idEmpresa);
        command.Parameters.AddWithValue("@IdMovimientoBancoEmisor", idMovimientoBancoEmisor);

        await connection.OpenAsync(cancellationToken);
        await command.ExecuteNonQueryAsync(cancellationToken);
    }

    private static CajaBancoResumenDto MapResumen(SqlDataReader reader)
    {
        return new CajaBancoResumenDto
        {
            IdMovimientoBanco = reader.GetInt32(reader.GetOrdinal("IdMovimientoBanco")),
            IdAsiento = reader.IsDBNull(reader.GetOrdinal("IdAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("IdAsiento")),
            NumeroAsiento = reader.IsDBNull(reader.GetOrdinal("NumeroAsiento")) ? null : reader.GetInt32(reader.GetOrdinal("NumeroAsiento")),
            NumeroMovimiento = reader.GetInt32(reader.GetOrdinal("NumeroMovimiento")),
            IdEmpresa = reader.GetInt32(reader.GetOrdinal("IdEmpresa")),
            IdBancoConfiguracionEmpresa = reader.GetInt32(reader.GetOrdinal("IdBancoConfiguracionEmpresa")),
            IdBanco = reader.GetInt32(reader.GetOrdinal("IdBanco")),
            NroCuentaCorriente = reader.GetString(reader.GetOrdinal("NroCuentaCorriente")),
            Titular = reader.IsDBNull(reader.GetOrdinal("Titular")) ? string.Empty : reader.GetString(reader.GetOrdinal("Titular")),
            CodigoBanco = reader.GetString(reader.GetOrdinal("CodigoBanco")),
            NombreBanco = reader.GetString(reader.GetOrdinal("NombreBanco")),
            CodigoMoneda = reader.IsDBNull(reader.GetOrdinal("CodigoMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("CodigoMoneda")),
            NombreMoneda = reader.IsDBNull(reader.GetOrdinal("NombreMoneda")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombreMoneda")),
            TipoMovimiento = reader.GetString(reader.GetOrdinal("TipoMovimiento")),
            IdOpeBancaria = reader.GetString(reader.GetOrdinal("IdOpeBancaria")),
            TipoOperacion = reader.IsDBNull(reader.GetOrdinal("TipoOperacion")) ? string.Empty : reader.GetString(reader.GetOrdinal("TipoOperacion")),
            FechaEmision = DateOnly.FromDateTime(reader.GetDateTime(reader.GetOrdinal("FechaEmision"))),
            IdPersona = reader.IsDBNull(reader.GetOrdinal("IdPersona")) ? null : reader.GetInt32(reader.GetOrdinal("IdPersona")),
            NumeroDocumentoPersona = reader.IsDBNull(reader.GetOrdinal("NumeroDocumentoPersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumentoPersona")),
            NombrePersona = reader.IsDBNull(reader.GetOrdinal("NombrePersona")) ? string.Empty : reader.GetString(reader.GetOrdinal("NombrePersona")),
            NumeroDocumento = reader.IsDBNull(reader.GetOrdinal("NumeroDocumento")) ? string.Empty : reader.GetString(reader.GetOrdinal("NumeroDocumento")),
            Glosa = reader.GetString(reader.GetOrdinal("Glosa")),
            ImporteTotal = reader.GetDecimal(reader.GetOrdinal("ImporteTotal")),
            Ingreso = reader.GetDecimal(reader.GetOrdinal("Ingreso")),
            Egreso = reader.GetDecimal(reader.GetOrdinal("Egreso")),
            Activo = reader.GetBoolean(reader.GetOrdinal("Activo"))
        };
    }
}
