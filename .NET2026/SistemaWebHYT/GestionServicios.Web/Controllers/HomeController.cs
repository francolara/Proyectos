using System.Security.Claims;
using GestionServicios.Application.Reporting;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Inicio con indicadores autorizados y cohorte de registro en fechas Perú.
// Firma: FRANCO LARA - 05/11/2026 | Filtros comerciales y bandeja histórica paginada del dashboard.
[Authorize]
public sealed class HomeController(IReportingRepository repository,TimeProvider clock):Controller
{
    public async Task<IActionResult> Index(DateOnly? from=null,DateOnly? to=null,string? client=null,string? store=null,int queuePage=1,CancellationToken ct=default)
    {
        if(!User.HasClaim(AccessService.PermissionClaim,"REPORTES.CONSULTAR"))return View();
        var first=from??ReportingDates.Today(clock).AddMonths(-1);var last=to??ReportingDates.Today(clock);
        client=string.IsNullOrWhiteSpace(client)?null:client.Trim();store=string.IsNullOrWhiteSpace(store)?null:store.Trim();
        if(!ModelState.IsValid||queuePage is <1 or >100000000||client?.Length>200||store?.Length>200||first>last||last==DateOnly.MaxValue)return BadRequest("Revisa las fechas del panel.");
        try{return View(new DashboardModel(first,last,await repository.DashboardAsync(User.FindFirstValue(ClaimTypes.NameIdentifier)!,ReportingDates.Start(first),ReportingDates.Start(last.AddDays(1)),ct,client,store,queuePage),client,store,queuePage));}
        catch(MaintenanceException error){ViewData["DashboardError"]=error.Message;return View();}
    }
    [AllowAnonymous, Route("/Error")]
    public IActionResult Error(){Response.StatusCode=500;return View("~/Views/Shared/Error.cshtml");}
}