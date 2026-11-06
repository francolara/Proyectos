using GestionServicios.Application.Maintenance;

sealed class TestTaxRateRepository : ITaxRateRepository
{
    public int Added { get; private set; }
    public Task<TaxRatePage> ListAsync(string actor, string? search, int page, CancellationToken ct)
        => Task.FromResult(new TaxRatePage([new(new DateOnly(2026,1,1), null, 0.18m, null, null, null, null, null)], 1, new DateOnly(2026,1,1)));
    public Task AddAsync(string actor, DateOnly from, decimal value, string reason, DateOnly? expectedLatest, CancellationToken ct)
    {
        if (expectedLatest != new DateOnly(2026,1,1)) throw new MaintenanceException("Las vigencias cambiaron. Actualiza el listado.");
        Added++; return Task.CompletedTask;
    }
}
