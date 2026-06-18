namespace SistemaAdministrativoWeb.Infrastructure.Data;

public sealed class PagedResult<T>
{
    public IReadOnlyCollection<T> Items { get; init; } = [];
    public int TotalRecords { get; init; }
    public int PageNumber { get; init; }
    public int PageSize { get; init; }
}
