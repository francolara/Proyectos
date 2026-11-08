using System.Security.Claims;
using GestionServicios.Application.Reporting;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.ViewModels.Reporting;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Reportes HTML paginados por cohorte y moneda, sin emisión fiscal.
// Firma: FRANCO LARA - 05/11/2026 | Conserva filtros del panel y entre páginas.
// Firma: FRANCO LARA - 06/11/2026 | Reportes integrales Presupuesto/OC, estados vigentes, pendientes y exportación HTML de todos los resultados. Rutas Index/Print separadas para que menú y enlaces de consulta nunca resuelvan a impresión.
[Authorize(Policy="perm:REPORTES.CONSULTAR")]
[Route("Reportes")]
public sealed class ReportsController(IReportingRepository repository,TimeProvider clock):Controller
{
    [HttpGet("")]
    public Task<IActionResult> Index(string? kind=null,DateOnly? from=null,DateOnly? to=null,int page=1,string? search=null,string? state=null,string? type=null,bool pendingGuide=false,bool pendingOrder=false,bool pendingConfirmation=false,bool pendingInvoice=false,CancellationToken ct=default)
        =>Query(kind,from,to,page,search,state,type,pendingGuide,pendingOrder,pendingConfirmation,pendingInvoice,false,ct);
    [HttpGet("Imprimir"), Authorize(Policy = "perm:REPORTES.IMPRIMIR")]
    public Task<IActionResult> Print(string? kind=null,DateOnly? from=null,DateOnly? to=null,int page=1,string? search=null,string? state=null,string? type=null,bool pendingGuide=false,bool pendingOrder=false,bool pendingConfirmation=false,bool pendingInvoice=false,CancellationToken ct=default)
        =>Query(kind,from,to,1,search,state,type,pendingGuide,pendingOrder,pendingConfirmation,pendingInvoice,true,ct);
    private async Task<IActionResult> Query(string? kind,DateOnly? from,DateOnly? to,int page,string? search,string? state,string? type,bool pendingGuide,bool pendingOrder,bool pendingConfirmation,bool pendingInvoice,bool print,CancellationToken ct)
    {
        kind??=User.HasClaim(AccessService.PermissionClaim,"PRESUPUESTOS.CONSULTAR")?"PRESUPUESTO":"OC_CLIENTE";
        var first=from??ReportingDates.Today(clock).AddMonths(-1);var last=to??ReportingDates.Today(clock);
        var states=kind=="PRESUPUESTO"?new[]{"BORRADOR","ENVIADO","OBSERVADO","APROBADO","RECHAZADO","CANCELADO"}:new[]{"REGISTRADA","TRABAJO_COMPLETADO","CERRADA_OPERATIVAMENTE","CANCELADA"};
        if(!ModelState.IsValid||!ReportKinds.All.ContainsKey(kind)||state is not null&&!states.Contains(state)||type is not (null or "INVITACION" or "EMERGENCIA")||search?.Length>200||page is <1 or >100000000||first>last||last==DateOnly.MaxValue||kind=="OC_CLIENTE"&&pendingOrder)return BadRequest("Revisa reporte, estado, fechas y filtros.");
        if(!User.HasClaim(AccessService.PermissionClaim,kind=="PRESUPUESTO"?"PRESUPUESTOS.CONSULTAR":"OC_CLIENTE.CONSULTAR")||pendingInvoice&&!User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR")||pendingConfirmation&&!User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.CONSULTAR")||pendingGuide&&!User.HasClaim(AccessService.PermissionClaim,"GUIAS.CONSULTAR")||pendingOrder&&!User.HasClaim(AccessService.PermissionClaim,"OC_CLIENTE.CONSULTAR"))return Forbid();
        var filters=new ReportFilter(search,state,type,pendingGuide,pendingOrder,pendingConfirmation,pendingInvoice,print);
        try{return View(print?"Print":"Index",new ReportModel(kind,first,last,await repository.ReportAsync(User.FindFirstValue(ClaimTypes.NameIdentifier)!,kind,ReportingDates.Start(first),ReportingDates.Start(last.AddDays(1)),print?1:page,ct,filters:filters),Filters:filters,Print:print));}
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
}
