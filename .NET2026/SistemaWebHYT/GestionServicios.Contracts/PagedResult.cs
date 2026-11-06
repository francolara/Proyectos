namespace GestionServicios.Contracts;

public sealed record PagedResult<T>(IReadOnlyList<T> Items, long TotalRecords, int Page)
{
    public const int PageSize = 20;
    public long TotalPages => TotalRecords / PageSize + (TotalRecords % PageSize == 0 ? 0 : 1);
}
