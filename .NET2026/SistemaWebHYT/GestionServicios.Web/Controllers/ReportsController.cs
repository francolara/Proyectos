using System.Security.Claims;
using GestionServicios.Application.Reporting;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.ViewModels.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Reportes HTML paginados por cohorte y moneda, sin emisión fiscal.
// Firma: FRANCO LARA - 05/11/2026 | Conserva cliente, tienda y seguimiento desde los accesos conciliados del panel y entre páginas.
[Authorize(Policy="perm:REPORTES.CONSULTAR")]
[Route("Reportes")]
public sealed class ReportsController(IReportingRepository repository,TimeProvider clock):Controller
{
    [HttpGet("")]
    public async Task<IActionResult> Index(string? kind=null,DateOnly? from=null,DateOnly? to=null,int page=1,string? client=null,string? store=null,string? situation=null,CancellationToken ct=default)
    {
        kind??=User.HasClaim(GestionServicios.Web.Security.AccessService.PermissionClaim,"PRESUPUESTOS.CONSULTAR")?"GANADOS":"OC_COMPLETADAS";
        var first=from??ReportingDates.Today(clock).AddMonths(-1);var last=to??ReportingDates.Today(clock);
        if(!ModelState.IsValid||(kind=="GANADOS"&&situation is not null)||(situation is not (null or "SIN_CONFIRMAR" or "CONFIRMADAS" or "LISTAS" or "FACTURADAS"))||client?.Length>200||store?.Length>200||!ReportKinds.All.ContainsKey(kind)||page is <1 or >100000000||first>last||last==DateOnly.MaxValue)return BadRequest("Revisa reporte, fechas y página.");
        if(!User.HasClaim(GestionServicios.Web.Security.AccessService.PermissionClaim,kind=="GANADOS"?"PRESUPUESTOS.CONSULTAR":"OC_CLIENTE.CONSULTAR"))return Forbid();
        try{return View(new ReportModel(kind,first,last,await repository.ReportAsync(User.FindFirstValue(ClaimTypes.NameIdentifier)!,kind,ReportingDates.Start(first),ReportingDates.Start(last.AddDays(1)),page,ct,client,store,situation),client,store,situation));}
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
}