using System.Text.Json;
namespace GestionServicios.Application.Budgets;

public sealed record BudgetVersion(long BudgetId,long RevisionId,string Number,int RevisionNumber,byte[] HeaderVersion,byte[] RevisionVersion);
public sealed record BudgetListItem(long BudgetId,string Number,string Description,string Type,long ClientId,long StoreId,long RevisionId,int RevisionNumber,
    string ClientName,string StoreName,string State,string Currency,decimal? Net,decimal? Tax,decimal? Total,DateOnly? ApprovedOn,
    DateTimeOffset CreatedAt,string CreatedBy,DateTimeOffset? UpdatedAt,string? UpdatedBy,long? PurchaseOrderId,string? CustomerOrderNumber,
    long? GuideId,string? GuideNumber,string? GuideState,long GuideAttachments,long? ExecutionId=null,bool Invoiced=false,string? InvoiceReferences=null);
public sealed record BudgetPage(IReadOnlyList<BudgetListItem> Rows,long Total,int Page,IReadOnlyDictionary<string,long>? StateCounts=null);
public sealed record BudgetAudit(long Id,string Action,string? PreviousState,string? NewState,DateTimeOffset RecordedAt,DateOnly? EventDate,string User,string? Reason,string? Changes,string? Entity=null)
{
    public string EntityLabel=>Entity switch{"Presupuesto"=>"Presupuesto","PresupuestoRevision"=>"Revisión de presupuesto","OrdenCompraCliente"=>"Orden de Compra - Cliente","GuiaConformidad"=>"Guía","ConfirmacionFacturacion"=>"Confirmación del cliente","Ejecucion"=>"Ejecución del trabajo","ArchivoAdjunto"=>"Archivo adjunto",_=>"Registro"};
}
public sealed record BudgetRecord(JsonElement Data,IReadOnlyList<BudgetAudit> Audit,long AuditTotal)
{
    public JsonElement Revision => Data.GetProperty("Revision");
    public string Text(string name) => Text(Data,name);
    public static string Text(JsonElement source,string name) => source.TryGetProperty(name,out var value) && value.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined) ? value.ToString() : "";
}
public sealed record TariffChoice(long TariffLineId,long MaterialId,string InternalCode,string? Position,string? SapCode,string Description,string Unit,
    string TariffCode,int Version,string TaxTreatment,string Currency,decimal? Price,string? ServiceType,string? NeedNumber,string? Zones);
public sealed record TariffChoicePage(IReadOnlyList<TariffChoice> Rows,long Total,int Page);
public interface IBudgetRepository
{
    Task<BudgetPage> ListAsync(string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,CancellationToken ct,string? client=null,string? store=null,string? type=null,bool pendingGuide=false,bool pendingOrder=false);
    Task<BudgetRecord?> GetAsync(string actor,long id,long? revisionId,int auditPage,CancellationToken ct);
    Task<BudgetVersion> SaveAsync(string actor,string payload,BudgetVersion? original,bool newRevision,CancellationToken ct);
    Task ChangeStateAsync(string actor,BudgetVersion original,string newState,DateOnly eventDate,string responsible,string? reason,CancellationToken ct);
    Task<TariffChoicePage> TariffsAsync(string actor,long clientId,DateOnly date,string currency,string? search,int page,CancellationToken ct,long? materialId=null,long? tariffLineId=null);
    Task<JsonElement> DefaultsAsync(string actor,DateOnly date,long? clientId,long? storeId,CancellationToken ct);
}
