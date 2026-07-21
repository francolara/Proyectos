using SistemaControlEspaciosDeportivosWeb.Configuration;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels.Legal;

public sealed class LegalDocumentViewModel
{
    public string Title { get; init; } = string.Empty;
    public string LastUpdated { get; init; } = string.Empty;
    public bool ShowDraftNotice { get; init; }
    public BusinessInformationOptions BusinessInformation { get; init; } = new();
}
