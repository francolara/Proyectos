using System.Text.Json;
using GestionServicios.Application.Budgets;
namespace GestionServicios.Application.Followup;

public sealed record DocumentKind(string Key,string Code,string Title,string Permission)
{
    public static IReadOnlyList<DocumentKind> All {get;}=Array.AsReadOnly(new[]{
        new DocumentKind("oc-cliente","OC","Orden de Compra - Cliente","OC_CLIENTE"),
        new DocumentKind("guias","GUIA","Guia (Acta de conformidad)","GUIAS"),
        new DocumentKind("confirmaciones","CONFIRMACION","Confirmación del cliente","CONFIRMACIONES"),
        new DocumentKind("ejecucion","EJECUCION","Ejecución del trabajo","EJECUCION")});
    public static DocumentKind Resolve(string key)=>All.FirstOrDefault(item=>item.Key==key)??throw new ArgumentException("Tipo de documento no válido.",nameof(key));
    public string Procedure(string operation)
    {
        if(!All.Contains(this)||operation is not ("Guardar" or "Obtener" or "Listar" or "CambiarEstado"))throw new ArgumentException("Operación documental no válida.");
        return "dbo.spu_Seguimiento"+Code+"_"+operation;
    }
}
public sealed record DocumentPage(IReadOnlyList<JsonElement> Rows,long Total,int Page);
public sealed record DocumentRecord(JsonElement Data,IReadOnlyList<BudgetAudit> Audit,long AuditTotal);
public sealed record DocumentSaved(long Id,long? FileId);
public interface IDocumentRepository
{
    Task<DocumentPage> ListAsync(DocumentKind kind,string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,bool? confirmed,CancellationToken ct,string? client=null,string? store=null);
    Task<DocumentRecord?> GetAsync(DocumentKind kind,string actor,long? id,long? parentId,int auditPage,CancellationToken ct);
    Task ChangeStateAsync(DocumentKind kind,string actor,long id,long parentId,byte[] parentVersion,byte[] originalVersion,Guid operation,string state,DateOnly eventDate,string responsible,string? reason,CancellationToken ct);
    Task<DocumentSaved> SaveAsync(DocumentKind kind,string actor,long parentId,byte[] parentVersion,string payload,Guid operationId,long? id,byte[]? originalVersion,CancellationToken ct);
}