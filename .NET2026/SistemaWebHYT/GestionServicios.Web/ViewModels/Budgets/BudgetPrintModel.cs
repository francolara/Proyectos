// Firma: FRANCO LARA - 06/11/2026 | Formato conforme al PDF, firma R2, subtítulo configurable y posición como código interno de material. Cabecera y cuadros ampliados; descripción del trabajo solo interna. Ancho transferido de Tienda a Centro SAP FICO. Arial y separación roja del subtítulo. Formato anterior adaptado al Excel con valores netos, sin filas IGV/total final y tienda descriptiva y contacto activo de tienda como Coordinador/Encargado.
using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.ViewModels.Budgets;

public sealed record BudgetPrintGroup(string Name,IReadOnlyList<JsonElement> Lines);
public sealed record BudgetPrintHistory(BudgetRecord Record,BudgetPrintPage Data);
public sealed class BudgetPrintModel
{
    public const string Version="HYT-2026.10";
    public required BudgetRecord Record {get;init;}
    public required string Format {get;init;}
    public required JsonElement Company {get;init;}
    public required string Logo {get;init;}
    public required string Css {get;init;}
    public required DateTimeOffset GeneratedAt {get;init;}
    public JsonElement Revision=>Record.Revision;
    public string Text(string name)=>BudgetRecord.Text(Revision,name);
    public string CompanyText(string name)=>BudgetRecord.Text(Company,name);
    public string Subtitle {get;set;}="";
    public string? Signature {get;set;}
    public string State=>Text("Estado");
    public bool IncludesTax=>Text("CodigoTipoIGV")=="CON_IGV";
    public string PriceTax=>IncludesTax?"CON IGV":"SIN IGV";
    public string Currency=>Text("CodigoMoneda");
    public string Symbol=>Text("SimboloMoneda");
    public string CurrencyLabel=>Symbol+" "+Currency;
    public static string Value(JsonElement source,string name)=>BudgetRecord.Text(source,name);
    public static string Money(decimal? value)=>value.HasValue?value.Value.ToString("N2",CultureInfo.GetCultureInfo("es-PE")):"Pendiente";
    public static decimal? Amount(JsonElement source,string name)=>source.TryGetProperty(name,out var value)&&value.ValueKind==JsonValueKind.Number?value.GetDecimal():null;
    public static string Date(string? value)=>DateOnly.TryParseExact(value,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out var date)?date.ToString("dd/MM/yyyy"):"—";
    public string Total(string field)=>Money(Amount(Revision,field));
    public decimal? LineAmount(JsonElement line)=>Amount(line,IncludesTax?"Total":"BaseImponible");
    public IReadOnlyList<BudgetPrintGroup> Groups=>Record.Data.GetProperty("Partidas").EnumerateArray().GroupBy(line=>Value(line,"Grupo")).Select(group=>new BudgetPrintGroup(group.Key,group.ToArray())).ToArray();
    public decimal? GroupAmount(BudgetPrintGroup group)=>group.Lines.Any(line=>!LineAmount(line).HasValue)?null:group.Lines.Sum(line=>LineAmount(line)!.Value);
    public decimal? NetUnitPrice(JsonElement line)
    {
        var price=Amount(line,"PrecioUnitario");
        return price.HasValue&&IncludesTax?BudgetAmounts.Round(price.Value/(1+Amount(Revision,"TasaIGVAplicada")!.Value)):price;
    }
    public decimal? NetGroupAmount(BudgetPrintGroup group)=>group.Lines.Any(line=>!Amount(line,"BaseImponible").HasValue)?null:group.Lines.Sum(line=>Amount(line,"BaseImponible")!.Value);
    public static BudgetPrintModel Create(BudgetRecord record,string format,string logo,string css,DateTimeOffset timestamp)
    {
        var snapshot=BudgetRecord.Text(record.Revision,"EmpresaSnapshotJson");
        if(string.IsNullOrWhiteSpace(snapshot))throw new MaintenanceException("Completa los datos de HYT y vuelve a guardar el borrador antes de generar el documento.");
        var company=JsonSerializer.Deserialize<JsonElement>(snapshot);
        if(company.ValueKind!=JsonValueKind.Object||string.IsNullOrEmpty(BudgetRecord.Text(company,"Ruc"))||string.IsNullOrEmpty(BudgetRecord.Text(company,"RazonSocial")))
            throw new MaintenanceException("Los datos corporativos conservados en esta revisión están incompletos.");
        if(format is not ("ANTERIOR" or "NUEVO"))throw new ArgumentException("Formato no válido.");
        return new(){Record=record,Format=format,Company=company,Logo=logo,Css=css,GeneratedAt=timestamp};
    }
}
