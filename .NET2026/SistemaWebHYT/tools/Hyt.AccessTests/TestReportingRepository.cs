using System.Text.Json;
using GestionServicios.Application.Reporting;
sealed class TestReportingRepository:IReportingRepository
{
    public int Page {get;private set;}
    public bool Empty {get;set;}
    public bool ShowBudgets {get;set;}=true;
    public bool ShowOrders {get;set;}=true;
    public int ReportRows {get;set;}=1;
    public int DashboardCalls {get;private set;}
    public Task<JsonElement> DashboardAsync(string actor,DateTimeOffset from,DateTimeOffset until,CancellationToken ct,string? client=null,string? store=null,int queuePage=1)
    {
        DashboardCalls++;
        var budgetCounts=new[]{3,4,8,6,9,7,5};var orderCounts=new[]{1,3,2,4,1,5,2};
        var movements=Empty?Array.Empty<object>():Enumerable.Range(0,7).Select(index=>(object)new{Fecha=until.UtcDateTime.AddHours(-5).Date.AddDays(index-7).ToString("yyyy-MM-dd"),Presupuestos=ShowBudgets?budgetCounts[index]:0,Ordenes=ShowOrders?orderCounts[index]:0,Confirmaciones=ShowOrders?1:0,Marcados=ShowOrders?1:0}).ToArray();
        return Task.FromResult(JsonSerializer.SerializeToElement(new{ActualizadoUTC="2026-11-05T17:30:00Z",VerPendientes=ShowBudgets&&ShowOrders,TotalPendientes=Empty?0:41,PendientesPeriodo=Empty?0:2,Pendientes=Empty?Array.Empty<object>():new object[]{new{Id=1,Numero="PPTO-0001",FechaRegistro="2026-10-01T17:00:00Z",TipoAtencion="EMERGENCIA",ClienteNombre="Cliente de prueba",TiendaNombre="Centro",SinOC=true,SinGuia=true,EmergenciaPorRegularizar=true,DiasPendiente=35}},VerPresupuestos=ShowBudgets,VerOC=ShowOrders,Presupuestos=Empty||!ShowBudgets?0:42,Ordenes=Empty||!ShowOrders?0:18,SinConfirmar=Empty||!ShowOrders?0:5,ConfirmadasSinMarcar=Empty||!ShowOrders?0:8,ListasParaMarcar=Empty||!ShowOrders?0:3,Facturadas=Empty||!ShowOrders?0:5,
            EstadosPresupuesto=Empty||!ShowBudgets?Array.Empty<object>():new object[]{new{Estado="BORRADOR",Cantidad=12},new{Estado="ENVIADO",Cantidad=18},new{Estado="APROBADO",Cantidad=12}},
            EstadosOC=Empty||!ShowOrders?Array.Empty<object>():new object[]{new{Estado="REGISTRADA",Cantidad=12},new{Estado="TRABAJO_COMPLETADO",Cantidad=4},new{Estado="CERRADA_OPERATIVAMENTE",Cantidad=2}},
            ImportesPresupuesto=Empty||!ShowBudgets?Array.Empty<object>():new object[]{new{CodigoMoneda="PEN",Importe=36000m,Cantidad=30,PreciosPendientes=0},new{CodigoMoneda="USD",Importe=1200m,Cantidad=12,PreciosPendientes=0}},
            ImportesOC=Empty||!ShowOrders?Array.Empty<object>():new object[]{new{CodigoMoneda="PEN",Importe=25000m,Cantidad=16},new{CodigoMoneda="USD",Importe=200m,Cantidad=2}},Movimientos=movements}));
    }
    public Task<ReportPage> ReportAsync(string actor,string kind,DateTimeOffset from,DateTimeOffset until,int page,CancellationToken ct,string? client=null,string? store=null,string? situation=null)
    {
        Page=page;
        var rows=Enumerable.Range(1,Math.Clamp(ReportRows,0,20)).Select(index=>JsonSerializer.SerializeToElement(new{Id=index,Numero=(kind=="GANADOS"?"PPTO-":"OC-")+index.ToString("D4")+"A",NumeroCliente="OC-0007B",FechaRegistro="2026-11-03T18:30:00Z",ClienteNombre="Cliente <script>prueba</script>",TiendaNombre="Centro",CodigoMoneda="PEN",Total=236m,Estado="APROBADO",Guia="ACTA-01A",Situacion="LISTA PARA MARCAR"})).ToArray();
        return Task.FromResult(new ReportPage(rows,ReportRows==0?0:41,ReportRows==0?[]:[new("PEN",9676m,41)],page));
    }
}