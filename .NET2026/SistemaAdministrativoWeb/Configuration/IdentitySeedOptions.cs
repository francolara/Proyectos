namespace SistemaAdministrativoWeb.Configuration;

public sealed class IdentitySeedOptions
{
    public const string SectionName = "IdentitySeed";

    public List<string> SuperAdminEmails { get; set; } = [];
}
