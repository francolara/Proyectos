using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.ViewModels.Maintenance;

// Firma: FRANCO LARA - 05/11/2026 | Captura simple de IGV, precios y estado con concurrencia propia.
public sealed class MaterialTariffInput
{
    public long? Id {get;set;}
    public string? RowVersion {get;set;}
    public string TaxCode {get;set;}="SIN_IGV";
    public string? PricePEN {get;set;}
    public string? PriceUSD {get;set;}
    public bool Active {get;set;}=true;
}
public sealed record MaterialTariffFormModel(long MaterialId,string MaterialLabel,MaterialTariffInput Input,RecordSnapshot? Record,bool ReadOnly);
