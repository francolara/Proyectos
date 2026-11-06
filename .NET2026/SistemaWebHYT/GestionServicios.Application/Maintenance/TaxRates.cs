namespace GestionServicios.Application.Maintenance;

public sealed record TaxRate(DateOnly From, DateOnly? Until, decimal Value, DateTimeOffset? CreatedAt, string? CreatedBy,
    DateTimeOffset? UpdatedAt, string? UpdatedBy, string? Reason);
public sealed record TaxRatePage(IReadOnlyList<TaxRate> Rows, long Total, DateOnly? Latest);
public interface ITaxRateRepository
{
    Task<TaxRatePage> ListAsync(string actor, string? search, int page, CancellationToken ct);
    Task AddAsync(string actor, DateOnly from, decimal value, string reason, DateOnly? expectedLatest, CancellationToken ct);
}
