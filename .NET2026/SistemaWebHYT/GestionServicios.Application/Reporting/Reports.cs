using System.Text.Json;
namespace GestionServicios.Application.Reporting;
public sealed record ReportAmount(string Currency,decimal? Amount,long Count);
public sealed record ReportPage(IReadOnlyList<JsonElement> Rows,long Total,IReadOnlyList<ReportAmount> Amounts,int Page);
public static class ReportKinds
{
    public static IReadOnlyDictionary<string,string> All {get;}=new Dictionary<string,string>{{"GANADOS","Presupuestos ganados"},{"OC_COMPLETADAS","OC completadas"},{"OC_CON_ACTA","OC con acta"},{"OC_PENDIENTES_FACTURA","OC pendientes de factura"}};
}
public static class ReportSituations
{
    public static IReadOnlyDictionary<string,string> All {get;}=new Dictionary<string,string>{{"SIN_CONFIRMAR","Sin confirmar"},{"CONFIRMADAS","Confirmadas sin marcado"},{"LISTAS","Listas para marcar"},{"FACTURADAS","Marcadas Facturado"}};
}
public interface IReportingRepository
{
    Task<JsonElement> DashboardAsync(string actor,DateTimeOffset from,DateTimeOffset until,CancellationToken ct,string? client=null,string? store=null,int queuePage=1);
    Task<ReportPage> ReportAsync(string actor,string kind,DateTimeOffset from,DateTimeOffset until,int page,CancellationToken ct,string? client=null,string? store=null,string? situation=null);
}