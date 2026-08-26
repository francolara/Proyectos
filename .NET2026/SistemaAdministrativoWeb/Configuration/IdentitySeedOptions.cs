namespace SistemaAdministrativoWeb.Configuration;

public sealed class IdentitySeedOptions
{
    public const string SectionName = "FRALSECONT_IdentitySeed";

    public List<string> SuperAdminEmails { get; set; } = [];
}
