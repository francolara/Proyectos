using System.Text.Json;

namespace GestionServicios.Application.Billing;

public sealed record FactoringLot(string Financiera, DateOnly FechaNegociacion, DateOnly? FechaDesembolso,
    decimal Porcentaje, string? Referencia, string? Observaciones, Guid OperacionId, IReadOnlyList<long> Facturas);

public interface IFactoringRepository
{
    Task<JsonElement> ReadFactoringAsync(string actor, long? id, DateOnly? from, DateOnly? to, string? search,
        int page, bool eligible, CancellationToken ct);
    Task<long> SaveFactoringAsync(string actor, FactoringLot lot, CancellationToken ct);
}
