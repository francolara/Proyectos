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
        // Firma: FRANCO LARA - 10/11/2026 | Barra negra compartida con acciones de impresión y CSV, incluidas impresiones conservadas.
        var bodyStart=document.Html.IndexOf("<body",StringComparison.OrdinalIgnoreCase);
        var toolbarAt=bodyStart<0?-1:document.Html.IndexOf('>',bodyStart);
        if(toolbarAt<0)return Content(document.Html,"text/html; charset=utf-8");

        var scriptUrl=System.Text.Encodings.Web.HtmlEncoder.Default.Encode(Url.Content("~/hyt-report-actions.js")!);
        var styleUrl=System.Text.Encodings.Web.HtmlEncoder.Default.Encode(Url.Content("~/hyt-report-toolbar.css")!);
        var toolbar=$"""
            <link rel="stylesheet" href="{styleUrl}" />
            <nav class="hyt-print-toolbar" aria-label="Acciones del presupuesto">
                <button type="button" data-report-print>Imprimir / Guardar PDF</button>
                <button type="button" data-report-csv>Exportar CSV</button>
            </nav>
            <script src="{scriptUrl}" defer></script>
            """;
        Response.Headers["Content-Security-Policy"]="default-src 'none'; img-src data:; style-src 'self' 'unsafe-inline'; script-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'self'";
        return Content(document.Html.Insert(toolbarAt+1,toolbar),"text/html; charset=utf-8");
    }
    [HttpGet("")]
    public async Task<IActionResult> Index(long budgetId,long revisionId,int page=1,CancellationToken ct=default)
    {
        if(budgetId<=0||revisionId<=0||page is <1 or >100000000)return BadRequest();
        var record=await budgets.GetAsync(Actor,budgetId,revisionId,1,ct);if(record is null)return NotFound();
        return View(new BudgetPrintHistory(record,await prints.ListAsync(Actor,revisionId,page,ct)));
    }
}
