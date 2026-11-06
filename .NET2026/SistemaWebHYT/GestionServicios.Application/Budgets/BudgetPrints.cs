namespace GestionServicios.Application.Budgets;

public sealed record BudgetPrint(long Id,long RevisionId,long BudgetId,string Format,string TemplateVersion,DateTimeOffset CreatedAt,string User,string Html,string Hash);
public sealed record BudgetPrintEntry(long Id,long RevisionId,long BudgetId,string Format,string TemplateVersion,DateTimeOffset CreatedAt,string User);
public sealed record BudgetPrintPage(IReadOnlyList<BudgetPrintEntry> Rows,long Total,int Page);
public interface IBudgetPrintRepository
{
    Task<long> SaveAsync(string actor,long revisionId,byte[] expectedVersion,string format,string templateVersion,string html,CancellationToken ct);
    Task<BudgetPrint?> GetAsync(string actor,long printId,CancellationToken ct);
    Task<BudgetPrintPage> ListAsync(string actor,long revisionId,int page,CancellationToken ct);
}
