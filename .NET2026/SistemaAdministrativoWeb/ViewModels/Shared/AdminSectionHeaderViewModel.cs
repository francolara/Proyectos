namespace SistemaAdministrativoWeb.ViewModels.Shared;

public sealed class AdminSectionHeaderViewModel
{
    public required string Id { get; init; }

    public required string Category { get; init; }

    public required string Title { get; init; }

    public required string Description { get; init; }

    public required string IconClass { get; init; }
}
