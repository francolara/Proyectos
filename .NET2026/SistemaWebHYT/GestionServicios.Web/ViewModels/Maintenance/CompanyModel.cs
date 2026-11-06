using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.ViewModels.Maintenance;
public sealed record CompanyModel(MaintenanceEditInput Input,RecordSnapshot? Record,bool CanEdit,int Page)
{
    public IReadOnlyDictionary<string,IReadOnlyList<CatalogOption>> Catalogs {get;init;}=new Dictionary<string,IReadOnlyList<CatalogOption>>();
}
