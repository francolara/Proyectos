namespace SistemaAdministrativoWeb.Infrastructure.Security;

public enum ModuleScope
{
    Unknown = 0,
    Account = 1,
    Company = 2
}

public sealed class ModuleAccessResult
{
    public bool IsAllowed { get; init; }
    public ModuleScope Scope { get; init; }
    public bool HasCompanyContext { get; init; }
    public string? Message { get; init; }
}
