using GestionServicios.Web.Security;
using System.Security.Claims;
using System.Text;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Services;
using GestionServicios.Web.ViewModels.Budgets;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 03/11/2026 | Generación y recuperación de impresiones HTML autorizadas y conservadas.
[Authorize(Policy="perm:PRESUPUESTOS.CONSULTAR")]
[Authorize(Policy="perm:REPORTES.CONSULTAR")]
[Route("Presupuestos/Impresiones")]
public sealed class BudgetReportsController(IBudgetRepository budgets,IBudgetPrintRepository prints,BudgetPrintRenderer renderer):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    [HttpPost("Generar"), Authorize(Policy = "perm:PRESUPUESTOS.IMPRIMIR")]
    public async Task<IActionResult> Generate(long budgetId,long revisionId,string format,string? version,CancellationToken ct)
    {
        if(budgetId<=0||revisionId<=0||format is not ("ANTERIOR" or "NUEVO"))return BadRequest("Formato no válido.");
        byte[] expected;try{expected=Convert.FromBase64String(version??"");}catch(FormatException){return BadRequest("Vuelve a abrir la revisión.");}
        if(expected.Length!=8)return BadRequest("Vuelve a abrir la revisión.");
        var record=await budgets.GetAsync(Actor,budgetId,revisionId,1,ct);if(record is null)return NotFound();
        if(!Convert.FromBase64String(BudgetRecord.Text(record.Revision,"VersionFila")).AsSpan().SequenceEqual(expected))return Conflict("La revisión cambió. Vuelve a consultarla.");
        try
        {
            var html=await renderer.RenderAsync(ControllerContext,record,format,ct);
            var id=await prints.SaveAsync(Actor,revisionId,expected,format,BudgetPrintModel.Version,html,ct);
            return RedirectToAction(nameof(Document),new {id});
        }
        catch(MaintenanceException error){TempData["Notice"]=error.Message;return RedirectToAction("Details","Budget",new {id=budgetId,revisionId});}
    }
    [HttpGet("{id:long}")]
    public async Task<IActionResult> Document(long id,bool download=false,CancellationToken ct=default)
    {
        if(download && !User.HasClaim(AccessService.PermissionClaim,"PRESUPUESTOS.DESCARGAR"))return Forbid();
        if(id<=0)return BadRequest();BudgetPrint? document;
        try{document=await prints.GetAsync(Actor,id,ct,download);}catch(MaintenanceException error){return Conflict(error.Message);}
        if(document is null)return NotFound();
        Response.Headers["Content-Security-Policy"]="default-src 'none'; img-src data:; style-src 'unsafe-inline'; script-src 'none'; base-uri 'none'; form-action 'none'; frame-ancestors 'self'";
        Response.Headers["X-Content-Type-Options"]="nosniff";
        if(download)return File(Encoding.UTF8.GetBytes(document.Html),"text/html; charset=utf-8",$"HYT-{document.BudgetId}-R{document.RevisionId}-{document.Format}.html");
        return Content(document.Html,"text/html; charset=utf-8");
    }
    [HttpGet("")]
    public async Task<IActionResult> Index(long budgetId,long revisionId,int page=1,CancellationToken ct=default)
    {
        if(budgetId<=0||revisionId<=0||page is <1 or >100000000)return BadRequest();
        var record=await budgets.GetAsync(Actor,budgetId,revisionId,1,ct);if(record is null)return NotFound();
        return View(new BudgetPrintHistory(record,await prints.ListAsync(Actor,revisionId,page,ct)));
    }
}
