using System.Text.Json;
namespace GestionServicios.Application.Reporting;
public sealed record ReportAmount(string Currency,decimal? Amount,long Count);
public sealed record ReportPage(IReadOnlyList<JsonElement> Rows,long Total,IReadOnlyList<ReportAmount> Amounts,int Page);
public static class ReportKinds
{
    public static IReadOnlyDictionary<string,string> All {get;}=new Dictionary<string,string>{{"PRESUPUESTO","Presupuesto"},{"OC_CLIENTE","Orden de Compra - Cliente"}};
}
public static class ReportSituations
{
    public static IReadOnlyDictionary<string,string> All {get;}=new Dictionary<string,string>{{"SIN_CONFIRMAR","Sin confirmar"},{"CONFIRMADAS","Confirmadas sin marcado"},{"LISTAS","Listas para marcar"},{"FACTURADAS","Marcadas Facturado"}};
}
public sealed record ReportFilter(string? Search=null,string? State=null,string? Type=null,bool PendingGuide=false,bool PendingOrder=false,bool PendingConfirmation=false,bool PendingInvoice=false,bool Export=false);
public interface IReportingRepository
{
    Task<JsonElement> DashboardAsync(string actor,DateTimeOffset from,DateTimeOffset until,CancellationToken ct,string? client=null,string? store=null,int queuePage=1);
    Task<ReportPage> ReportAsync(string actor,string kind,DateTimeOffset from,DateTimeOffset until,int page,CancellationToken ct,string? client=null,string? store=null,string? situation=null,ReportFilter? filters=null);
}