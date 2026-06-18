namespace SistemaAdministrativoWeb.Configuration;

public sealed class IdentityBehaviorSettings
{
    public const string SectionName = "IdentityBehavior";

    public bool RequireConfirmedAccount { get; set; }
    public bool AutoConfirmEmail { get; set; } = true;
}
