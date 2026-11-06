using System.Data;
using GestionServicios.Application.Maintenance;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Maintenance;

public sealed class SqlTaxRateRepository(SqlConnectionFactory connections) : ITaxRateRepository
{
    public async Task<TaxRatePage> ListAsync(string actor, string? search, int page, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = new SqlCommand("dbo.spu_MantTasaIGV_Listar", connection) { CommandType = CommandType.StoredProcedure };
        command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
        command.Parameters.Add("@Busqueda", SqlDbType.NVarChar, 200).Value = (object?)search ?? DBNull.Value;
        command.Parameters.Add("@Pagina", SqlDbType.Int).Value = page;
        await using var reader = await command.ExecuteReaderAsync(ct);
        var rows = new List<TaxRate>();
        DateTimeOffset? Timestamp(int index) => reader.IsDBNull(index) ? null : new DateTimeOffset(DateTime.SpecifyKind(reader.GetDateTime(index), DateTimeKind.Utc));
        string? Text(int index) => reader.IsDBNull(index) ? null : reader.GetString(index);
        while (await reader.ReadAsync(ct)) rows.Add(new(DateOnly.FromDateTime(reader.GetDateTime(0)), reader.IsDBNull(1) ? null : DateOnly.FromDateTime(reader.GetDateTime(1)),
            reader.GetDecimal(2), Timestamp(3), Text(4), Timestamp(5), Text(6), Text(7)));
        await reader.NextResultAsync(ct); await reader.ReadAsync(ct);
        var total = reader.GetInt64(0);
        var latest = reader.IsDBNull(1) ? (DateOnly?)null : DateOnly.FromDateTime(reader.GetDateTime(1));
        return new(rows, total, latest);
    }

    public async Task AddAsync(string actor, DateOnly from, decimal value, string reason, DateOnly? expectedLatest, CancellationToken ct)
    {
        try
        {
            await using var connection = connections.CreateConnection();
            await connection.OpenAsync(ct);
            await using var command = new SqlCommand("dbo.spu_MantTasaIGV_Agregar", connection) { CommandType = CommandType.StoredProcedure };
            command.Parameters.Add("@UsuarioId", SqlDbType.NVarChar, 128).Value = actor;
            command.Parameters.Add("@CorrelacionId", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
            command.Parameters.Add("@FechaDesde", SqlDbType.Date).Value = from.ToDateTime(TimeOnly.MinValue);
            command.Parameters.Add("@UltimaFechaEsperada", SqlDbType.Date).Value = expectedLatest.HasValue ? expectedLatest.Value.ToDateTime(TimeOnly.MinValue) : DBNull.Value;
            var parameter = command.Parameters.Add("@Valor", SqlDbType.Decimal); parameter.Precision = 9; parameter.Scale = 6; parameter.Value = value;
            command.Parameters.Add("@Motivo", SqlDbType.NVarChar, 1000).Value = reason;
            await command.ExecuteNonQueryAsync(ct);
        }
        catch (SqlException error) when (error.Message.Contains("[HYT_", StringComparison.Ordinal))
        { throw new MaintenanceException(error.Message[(error.Message.IndexOf(']') + 1)..].Trim()); }
    }
}
