namespace SistemaAdministrativoWeb.Configuration;

public sealed class LegalDocumentsOptions
{
    public const string SectionName = "FRALSECONT_LegalDocuments";

    public string TermsLastUpdated { get; set; } = "10/07/2026";
    public string PrivacyLastUpdated { get; set; } = "10/07/2026";
    public string CancellationsLastUpdated { get; set; } = "10/07/2026";
}
