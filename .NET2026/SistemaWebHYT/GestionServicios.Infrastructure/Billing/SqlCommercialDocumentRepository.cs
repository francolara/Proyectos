using System.Data;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Billing;

public sealed class SqlCommercialDocumentRepository(SqlConnectionFactory connections) : ICommercialDocumentRepository, IFactoringRepository
{
    private async Task<T> ExecuteAsync<T>(string operation, string actor, (string, SqlDbType, object?)[] values,
        Func<object?, T> read, CancellationToken ct)
    {
        try
        {
            await using var connection = connections.CreateConnection();
            await connection.OpenAsync(ct);
            await using var command = new SqlCommand("dbo.spu_ComprobanteHYT_" + operation, connection)
                { CommandType = CommandType.StoredProcedure };
            command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
            foreach (var (key, type, value) in values)
            {
                var parameter = command.Parameters.Add(key, type);
                if (type == SqlDbType.NVarChar) parameter.Size = -1;
                parameter.Value = value is DateOnly date ? date.ToDateTime(TimeOnly.MinValue) : value ?? DBNull.Value;
                if (type == SqlDbType.Decimal) { parameter.Precision = 18; parameter.Scale = 2; }
            }
            return read(await command.ExecuteScalarAsync(ct));
        }
        catch (SqlException ex) when (ex.Number is 2601 or 2627)
        {
            throw new MaintenanceException("El comprobante o la referencia ya se registraron.");
        }
        catch (SqlException ex) when (ex.Number is 8115 or 8152 or 2628)
        {
            throw new MaintenanceException("Un importe o texto supera la longitud permitida.");
        }
        catch (SqlException ex) when (ex.Message.StartsWith("[HYT_", StringComparison.Ordinal))
        {
            throw new MaintenanceException(ex.Message[(ex.Message.IndexOf(']') + 1)..].Trim());
        }
    }

    public Task<JsonElement> ReadAsync(string actor, bool sales, long? id, DateOnly? from, DateOnly? to,
        string? type, string? search, int page, long? clientId, long? budgetId, CancellationToken ct, bool withoutFactoring = false, bool? hasFactoring = null,
        IReadOnlyList<string>? materialCodes = null, bool orderLookup = false, bool invoiceLookup = false)
        => ExecuteAsync("Consultar", actor, [
            ("@EsVenta", SqlDbType.Bit, sales), ("@Id", SqlDbType.BigInt, id),
            ("@Desde", SqlDbType.Date, from), ("@Hasta", SqlDbType.Date, to),
            ("@Tipo", SqlDbType.Char, type), ("@Busqueda", SqlDbType.NVarChar, search), ("@Pagina", SqlDbType.Int, page),
            ("@ClienteId", SqlDbType.BigInt, clientId), ("@PresupuestoId", SqlDbType.BigInt, budgetId),
            ("@SinFactoring", SqlDbType.Bit, withoutFactoring), ("@ConFactoring", SqlDbType.Bit, hasFactoring),
            ("@DetalleMateriales", SqlDbType.NVarChar, materialCodes is null ? null : JsonSerializer.Serialize(materialCodes.Select(code => new { Codigo = code }))),
            ("@BuscarOrdenes", SqlDbType.Bit, orderLookup), ("@BuscarFacturas", SqlDbType.Bit, invoiceLookup)
        ], value => JsonSerializer.Deserialize<JsonElement>((string)value!), ct);

    public Task<JsonElement> ReadFactoringAsync(string actor, long? id, DateOnly? from, DateOnly? to, string? search,
        int page, bool eligible, CancellationToken ct)
        => ExecuteAsync("FactoringConsultar", actor, [
            ("@Id", SqlDbType.BigInt, id), ("@Desde", SqlDbType.Date, from), ("@Hasta", SqlDbType.Date, to),
            ("@Busqueda", SqlDbType.NVarChar, search), ("@Pagina", SqlDbType.Int, page), ("@Elegibles", SqlDbType.Bit, eligible)
        ], value => JsonSerializer.Deserialize<JsonElement>((string)value!), ct);

    public Task<long> SaveFactoringAsync(string actor, FactoringLot lot, CancellationToken ct)
        => ExecuteAsync("FactoringGuardar", actor, [("@Datos", SqlDbType.NVarChar, JsonSerializer.Serialize(lot))], Convert.ToInt64, ct);

    public Task<long> SaveAsync(string actor, bool sales, CommercialDocument document, long? id, CancellationToken ct)
        => ExecuteAsync("Guardar", actor, [("@EsVenta", SqlDbType.Bit, sales), ("@Id", SqlDbType.BigInt, id),
            ("@Datos", SqlDbType.NVarChar, JsonSerializer.Serialize(document))], Convert.ToInt64, ct);

    public async Task ImportBatchAsync(string actor, bool sales,
        IReadOnlyList<(string File, CommercialDocument Document)> files, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var transaction = (SqlTransaction)await connection.BeginTransactionAsync(ct);
        await using var command = new SqlCommand("dbo.spu_ComprobanteHYT_Guardar", connection, transaction)
            { CommandType = CommandType.StoredProcedure };
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
        command.Parameters.Add("@EsVenta", SqlDbType.Bit).Value = sales;
        var payload = command.Parameters.Add("@Datos", SqlDbType.NVarChar, -1);
        foreach (var (file, document) in files)
        {
            payload.Value = JsonSerializer.Serialize(document);
            try { await command.ExecuteScalarAsync(ct); }
            catch (SqlException ex)
            {
                var reason = ex.Message.StartsWith("[HYT_", StringComparison.Ordinal)
                    ? ex.Message[(ex.Message.IndexOf(']') + 1)..].Trim()
                    : "SQL Server rechazó el registro.";
                throw new MaintenanceException($"{file}: {reason}");
            }
        }
        await transaction.CommitAsync(ct);
    }

    public Task LinkAsync(string actor, bool sales, long id, IReadOnlyList<long> ids, long? budgetId,
        string? description, decimal? amount, CancellationToken ct, Guid? operationId = null)
        => ExecuteAsync("Vincular", actor, [("@EsVenta", SqlDbType.Bit, sales), ("@Id", SqlDbType.BigInt, id),
            ("@Ids", SqlDbType.NVarChar, JsonSerializer.Serialize(ids)), ("@PresupuestoId", SqlDbType.BigInt, budgetId),
            ("@OperacionId", SqlDbType.UniqueIdentifier, operationId), ("@Descripcion", SqlDbType.NVarChar, description), ("@Importe", SqlDbType.Decimal, amount)], Convert.ToInt64, ct);

    public Task LinkOrderAsync(string actor, long id, long? orderId, CancellationToken ct)
        => ExecuteAsync("Vincular", actor, [
            ("@EsVenta", SqlDbType.Bit, true), ("@Id", SqlDbType.BigInt, id),
            ("@ActualizarOrdenCompra", SqlDbType.Bit, true), ("@OrdenCompraId", SqlDbType.BigInt, orderId)
        ], Convert.ToInt64, ct);

    public Task<JsonElement> ReadWithholdingsAsync(string actor, string state, string? supplier, string? search, int page, CancellationToken ct)
        => ExecuteAsync("DetraccionesConsultar", actor, [("@Estado", SqlDbType.NVarChar, state),
            ("@Proveedor", SqlDbType.NVarChar, supplier), ("@Busqueda", SqlDbType.NVarChar, search), ("@Pagina", SqlDbType.Int, page)],
            value => JsonSerializer.Deserialize<JsonElement>((string)value!), ct);

    public Task PayAsync(string actor, long id, bool withholding, DateOnly date, decimal amount, string reference, CancellationToken ct)
        => ExecuteAsync("Pago", actor, [("@Id", SqlDbType.BigInt, id), ("@EsDetraccion", SqlDbType.Bit, withholding),
            ("@Fecha", SqlDbType.Date, date), ("@Importe", SqlDbType.Decimal, amount),
            ("@Referencia", SqlDbType.NVarChar, reference)], Convert.ToInt64, ct);
}
