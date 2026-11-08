using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Followup;
using GestionServicios.Web.ViewModels.Shared;
namespace GestionServicios.Web.ViewModels.Followup;
public sealed record DocumentListModel(DocumentKind Kind,GridModel Grid,FilterModel Filter,IReadOnlyList<StateKpi>? Kpis=null);
public sealed record DocumentDetailsModel(DocumentKind Kind,DocumentRecord Record,int AuditPage)
{
    public string Text(string name)=>BudgetRecord.Text(Record.Data,name);
    public JsonElement Document=>Record.Data.GetProperty("Documento");
    public static string Display(JsonElement value,string name)
    {
        var text=BudgetRecord.Text(value,name);if(text.Length==0)return "—";
        if(name.StartsWith("Fecha",StringComparison.Ordinal)&&DateTimeOffset.TryParse(text,CultureInfo.InvariantCulture,DateTimeStyles.AssumeUniversal,out var date))
            return name is "FechaRegistro" or "FechaActualizacion"?PeruDate.Display(date):date.ToString("dd/MM/yyyy");
        if(name=="Importe"&&decimal.TryParse(text,CultureInfo.InvariantCulture,out var money))return money.ToString("N2",CultureInfo.GetCultureInfo("es-PE"));
        return text;
    }
    public IReadOnlyList<(string Name,string Label)> Fields=>Kind.Code switch
    {
        "OC"=>[("NumeroRegistroInterno","Registro interno"),("NumeroOCCliente","OC del cliente"),("FechaEmisionOC","Emisión"),("FechaRecepcionOC","Recepción"),("CodigoMoneda","Moneda"),("Importe","Importe"),("Observacion","Observación")],
        "GUIA"=>[("NumeroActa","Número de guía"),("FechaActa","Fecha de guía"),("FechaAceptacion","Aceptación"),("RepresentanteCliente","Representante del cliente"),("Observacion","Observación")],
        "CONFIRMACION"=>[("NumeroConfirmacion","Número de confirmación"),("FechaConfirmacion","Fecha de confirmación"),("RepresentanteCliente","Representante del cliente"),("Observacion","Observación")],
        _=>[("Responsable","Responsable"),("FechaProgramada","Programación"),("FechaInicioReal","Inicio real"),("FechaTerminoReal","Término real"),("Descripcion","Trabajo realizado")]
    };
}
public sealed class DocumentInputModel
{
    public long ParentId {get;set;}
    public long? Id {get;set;}
    public string ParentVersion {get;set;}="";
    public string? OriginalVersion {get;set;}
    public Guid OperationId {get;set;}
    public Dictionary<string,string?> Fields {get;set;}=[];
    public IFormFile? Attachment {get;set;}
    public List<long> RemovedFiles {get;set;}=[];
    public string? MotivoCambio {get;set;}
    public string? ReturnUrl {get;set;}
}
public sealed record DocumentFormModel(DocumentKind Kind,DocumentInputModel Input,DocumentRecord Record,bool Popup=false,bool ReadOnly=false);