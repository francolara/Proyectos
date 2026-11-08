using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Reporting;
using GestionServicios.Application.Budgets;
using GestionServicios.Web.ViewModels.Budgets;
namespace GestionServicios.Web.ViewModels.Reporting;
public static class ReportDisplay
{
    public static IReadOnlyDictionary<string,string> States(string kind)=>kind=="PRESUPUESTO"?new Dictionary<string,string>{{"BORRADOR","Registrado"},{"ENVIADO","Enviado"},{"OBSERVADO","Observado"},{"APROBADO","Aceptada"},{"RECHAZADO","Rechazada"},{"CANCELADO","Cancelado"}}:new Dictionary<string,string>{{"REGISTRADA","Registrada"},{"TRABAJO_COMPLETADO","Completado"},{"CERRADA_OPERATIVAMENTE","Cerrada operativamente"},{"CANCELADA","Cancelada"}};
    public static string State(string code,bool budget=false)=>budget?BudgetStates.Label(code):States("OC_CLIENTE").GetValueOrDefault(code,code);
    public static string Date(JsonElement row,string key)=>DateTimeOffset.TryParse(BudgetRecord.Text(row,key),CultureInfo.InvariantCulture,DateTimeStyles.AssumeUniversal,out var date)?TimeZoneInfo.ConvertTime(date,ReportingDates.Lima).ToString("dd/MM/yyyy HH:mm"):"—";
    public static string BusinessDate(JsonElement row,string key)=>DateOnly.TryParse(BudgetRecord.Text(row,key).Split('T')[0],out var date)?date.ToString("dd/MM/yyyy"):"—";
    public static Dictionary<string,string> Query(ReportModel model,int? page=null)
    {
        var f=model.Filters??new ReportFilter();return new(){["kind"]=model.Kind,["from"]=model.From.ToString("yyyy-MM-dd"),["to"]=model.To.ToString("yyyy-MM-dd"),["search"]=f.Search??"",["state"]=f.State??"",["type"]=f.Type??"",["pendingGuide"]=f.PendingGuide.ToString().ToLowerInvariant(),["pendingOrder"]=f.PendingOrder.ToString().ToLowerInvariant(),["pendingConfirmation"]=f.PendingConfirmation.ToString().ToLowerInvariant(),["pendingInvoice"]=f.PendingInvoice.ToString().ToLowerInvariant(),["page"]=(page??1).ToString(CultureInfo.InvariantCulture)};
    }
    public static string QueryString(ReportModel model,int? page=null)=>string.Join("&",Query(model,page).Select(pair=>Uri.EscapeDataString(pair.Key)+"="+Uri.EscapeDataString(pair.Value)));
    public static string Filters(ReportModel model)
    {
        var f=model.Filters??new ReportFilter();var parts=new List<string>{"Registro: "+model.From.ToString("dd/MM/yyyy")+" al "+model.To.ToString("dd/MM/yyyy"),"Estado: "+(f.State is null?"Todos":States(model.Kind).GetValueOrDefault(f.State,f.State)),"Tipo: "+(f.Type=="INVITACION"?"Invitación":f.Type=="EMERGENCIA"?"Emergencia":"Todos"),"Buscar: "+(string.IsNullOrWhiteSpace(f.Search)?"Todos":f.Search)};
        if(f.PendingGuide)parts.Add("Pendiente de guía");if(f.PendingOrder)parts.Add("Pendiente de OC del cliente");if(f.PendingConfirmation)parts.Add("Pendiente de confirmación del cliente");if(f.PendingInvoice)parts.Add("Pendiente de facturar");return string.Join(" · ",parts);
    }
}
