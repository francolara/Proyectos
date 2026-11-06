using GestionServicios.Application.Maintenance;
using GestionServicios.Web.ViewModels.Shared;

namespace GestionServicios.Web.ViewModels.Maintenance;

// Firma: FRANCO LARA - 05/11/2026 | Códigos internos del servidor y tarifas contextuales y lista editable de enlaces de documentos por cliente.
public sealed class MaintenanceListModel
{
    public required EntityDefinition Entity { get; init; }
    public required GridModel Grid { get; init; }
    public string Search { get; init; } = "";
    public long? ParentId { get; init; }
    public bool CanEdit { get; init; }
}
public sealed class MaintenanceEditInput
{
    public bool Compact {get;set;}
    public long? MaterialContextId {get;set;}
    public long? Id { get; set; }
    public string? RowVersion { get; set; }
    public Dictionary<string, string?> Fields { get; set; } = [];
    public List<int> ZoneIds { get; set; } = [];
    public List<ClientDocumentLink> DocumentLinks { get; set; } = [];
    public string? ReturnUrl { get; set; }
}
public sealed class MaintenanceFormModel
{
    public required EntityDefinition Entity { get; init; }
    public required MaintenanceEditInput Input { get; init; }
    public RecordSnapshot? Record { get; init; }
    public IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>> Catalogs { get; init; } = new Dictionary<string, IReadOnlyList<CatalogOption>>();
    public bool ReadOnly { get; init; }
    public bool CanEdit { get; init; }
    public int AuditPage { get; init; } = 1;
    public string? MaterialContextLabel {get;init;}
    public MaterialTariffsModel? MaterialTariffs {get;init;}
    public string? Error { get; init; }
}

public sealed record MaterialTariffsModel(long MaterialId,string MaterialLabel,MaterialTariffPage Data,bool CanEdit);
