namespace GestionServicios.Application.Payroll;

public sealed record Supplier(
    long Id,
    string TaxId,
    string Name,
    string? Bank,
    string? MaskedAccount,
    string? Account,
    bool Active,
    byte[] Version,
    long Total,
    string? Cci = null);

public sealed record SupplierCost(
    long Id,
    long BudgetId,
    string BudgetNumber,
    long SupplierId,
    string SupplierName,
    string Description,
    string Currency,
    string? Document,
    decimal Amount,
    decimal Paid,
    decimal Balance,
    string Status,
    string? Bank,
    string? Account,
    long Total,
    DateTime? CostDate = null,
    string? PurchaseInvoice = null);

public sealed record SupplierCostFilters(DateOnly? From, DateOnly? To, string Status, string? Search);

public sealed record SupplierPayment(
    long Id,
    decimal Amount,
    DateTime Date,
    string Reference,
    string Currency,
    long Total,
    string? Frequency = null,
    DateTime? PeriodFrom = null,
    DateTime? PeriodTo = null);

public sealed record SupplierInput(
    long? Id,
    string TaxId,
    string Name,
    bool Active,
    byte[]? Version,
    bool UpdateAccount,
    string? Bank,
    string? Account,
    string? Cci = null,
    bool UpdateCci = false);

public sealed record SupplierCostInput(
    Guid OperationId,
    long BudgetId,
    long SupplierId,
    string Description,
    decimal Amount,
    string Currency,
    string? Document);

public sealed record SupplierPaymentInput(
    Guid OperationId,
    long CostId,
    decimal Amount,
    DateOnly Date,
    string Reference,
    string? Frequency = null,
    DateOnly? PeriodFrom = null,
    DateOnly? PeriodTo = null);

public interface ISupplierPaymentsRepository
{
    Task<IReadOnlyList<Supplier>> SuppliersAsync(string actor, long? supplierId, string? search, bool activeOnly, int page, CancellationToken ct);
    Task<IReadOnlyList<SupplierCost>> CostsAsync(string actor, long? budgetId, long? supplierId, long? costId, bool pendingOnly, int page, CancellationToken ct, SupplierCostFilters? filters = null);
    Task<IReadOnlyList<SupplierPayment>> PaymentsAsync(string actor, long costId, int page, CancellationToken ct);
    Task<long> SaveSupplierAsync(string actor, SupplierInput input, CancellationToken ct);
    Task<long> SaveCostAsync(string actor, SupplierCostInput input, CancellationToken ct);
    Task<long> SavePaymentAsync(string actor, SupplierPaymentInput input, CancellationToken ct);
}
