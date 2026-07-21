namespace SistemaControlEspaciosDeportivosWeb.Configuration;

public sealed class LegalDocumentsOptions
{
    public const string SectionName = "LegalDocuments";

    public string TermsLastUpdated { get; set; } = "20/07/2026";
    public string PrivacyLastUpdated { get; set; } = "20/07/2026";
    public string CancellationsLastUpdated { get; set; } = "20/07/2026";
}
