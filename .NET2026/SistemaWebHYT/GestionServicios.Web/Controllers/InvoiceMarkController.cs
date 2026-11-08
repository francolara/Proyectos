// Firma: FRANCO LARA - 06/11/2026 | Consulta informativa del comprobante externo en pop-up sin edición.
using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Billing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Referencia externa Facturado, corrección/retiro auditados y captura popup.
[Authorize(Policy="perm:FACTURADO.CONSULTAR")]
[Authorize(Policy="perm:OC_CLIENTE.CONSULTAR")]
[Route("OCCliente/{orderId:long}/Facturado")]
public sealed class InvoiceMarkController(IInvoiceMarkRepository repository):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit=>User.HasClaim(AccessService.PermissionClaim,"FACTURADO.EDITAR");
    private string Back(long orderId,string? url)=>Url.IsLocalUrl(url)?url!:Url.Action("Details","Followup",new{kind="oc-cliente",id=orderId})!;
    private IActionResult Form(InvoiceMarkModel model)=>model.Popup?PartialView("_Form",model):View("Index",model);
    private static byte[]? Version(string? text){try{var bytes=Convert.FromBase64String(text??"");return bytes.Length==8?bytes:null;}catch(FormatException){return null;}}
    [HttpGet("")]
    public async Task<IActionResult> Index(long orderId,bool popup=false,string? returnUrl=null,int page=1,CancellationToken ct=default)
    {
        if(orderId<=0||page is <1 or >100000000)return BadRequest();
        try
        {
            var data=await repository.GetAsync(Actor,orderId,page,ct);if(data is null)return NotFound();var record=data.Value;
            var input=new InvoiceMarkInput{OrderId=orderId,OrderVersion=BudgetRecord.Text(record,"VersionFila"),OperationId=Guid.NewGuid(),ReturnUrl=Back(orderId,returnUrl)};
            if(record.TryGetProperty("Referencia",out var reference)&&reference.ValueKind==JsonValueKind.Object)
            {
                input.ReferenceId=reference.GetProperty("ReferenciaId").GetInt64();input.ReferenceVersion=BudgetRecord.Text(reference,"VersionFila");
                input.Date=DateOnly.Parse(BudgetRecord.Text(reference,"FechaDocumento")[..10],System.Globalization.CultureInfo.InvariantCulture);input.Type=BudgetRecord.Text(reference,"CodigoTipoDocumento");input.Series=BudgetRecord.Text(reference,"Serie");input.Number=BudgetRecord.Text(reference,"NumeroDocumento");
            }
            return Form(new(input,record,CanEdit,popup,page));
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpGet("Consulta")]
    public async Task<IActionResult> Summary(long orderId,CancellationToken ct=default)
    {
        if(orderId<=0)return BadRequest();
        try
        {
            var data=await repository.GetAsync(Actor,orderId,1,ct);
            if(data is null||!data.Value.TryGetProperty("Referencia",out var reference)||reference.ValueKind!=JsonValueKind.Object)return NotFound();
            return PartialView("_Summary",reference);
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpPost("")]
    [Authorize(Policy="perm:FACTURADO.EDITAR")]
    public async Task<IActionResult> Save(long orderId,InvoiceMarkInput input,bool popup=false,CancellationToken ct=default)
    {
        var version=Version(input.OrderVersion);var referenceVersion=Version(input.ReferenceVersion);
        if(input.OrderId!=orderId||orderId<=0||version is null||input.OperationId==Guid.Empty||input.ReferenceId.HasValue&&referenceVersion is null)return BadRequest("Vuelve a abrir la OC.");
        input.ReturnUrl=Back(orderId,input.ReturnUrl);
        if(!input.Withdraw&&(input.Date is null||input.Type is not("01" or "03")||string.IsNullOrWhiteSpace(input.Series)||string.IsNullOrWhiteSpace(input.Number)))ModelState.AddModelError("","Completa fecha, tipo, serie y número del comprobante externo.");
        if((input.ReferenceId.HasValue||input.Withdraw)&&string.IsNullOrWhiteSpace(input.Reason))ModelState.AddModelError("","Indica el motivo de la corrección o retiro.");
        if(ModelState.IsValid)
        {
            try
            {
                await repository.SaveAsync(Actor,new(orderId,version,input.OperationId,input.Date,input.Type,input.Series,input.Number,input.ReferenceId,referenceVersion,input.Withdraw,input.Reason),ct);
                if(popup)return Json(new{redirect=input.ReturnUrl});return LocalRedirect(input.ReturnUrl);
            }
            catch(MaintenanceException error){ModelState.AddModelError("",error.Message);}
        }
        var record=await repository.GetAsync(Actor,orderId,1,ct);if(record is null)return NotFound();Response.StatusCode=422;return Form(new(input,record.Value,true,popup,1));
    }
}