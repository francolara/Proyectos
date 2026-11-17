using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.ViewModels.Maintenance;
// Firma: FRANCO LARA - 17/11/2026 | Catálogo de firmas con nombre, estado y predeterminada.
public sealed record CompanyModel(MaintenanceEditInput Input,RecordSnapshot? Record,bool CanEdit,int Page)
{
    public IReadOnlyList<BudgetSignatureOption> Signatures {get;init;}=[];
    public IReadOnlyDictionary<string,IReadOnlyList<CatalogOption>> Catalogs {get;init;}=new Dictionary<string,IReadOnlyList<CatalogOption>>();
}
