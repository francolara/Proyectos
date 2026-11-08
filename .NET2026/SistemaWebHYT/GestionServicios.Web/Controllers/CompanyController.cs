// Firma: FRANCO LARA - 06/11/2026 | Ubigeo HYT, firma de presupuesto JPG/PNG privada en R2 y subtítulo de cotización configurable. Firma almacenada en ConfigHYT/Firma.
using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Web.Services;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 03/11/2026 | Configuración única de HYT con concurrencia y auditoría.
[Authorize(Policy="perm:CONFIGURACION.CONSULTAR")]
[Route("Configuracion/HYT")]
public sealed class CompanyController(ICompanySettingsRepository repository,IMaintenanceRepository maintenance,IDocumentFileStore files):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit=>User.HasClaim(AccessService.PermissionClaim,"CONFIGURACION.EDITAR");
    [HttpGet("")]
    public async Task<IActionResult> Index(int page=1,CancellationToken ct=default)
    {
        if(page is <1 or >100000000)return BadRequest();var record=await repository.GetAsync(Actor,page,ct);
        var input=new MaintenanceEditInput{Id=record is null?null:1,RowVersion=record?.Text("VersionFila")};
        foreach(var field in CompanySettings.Definition.Fields)input.Fields[field.Name]=record?.Text(field.Name);
        if(record is null)input.Fields["SubtituloCotizacion"]=CompanySettings.DefaultSubtitle;
        return View(new CompanyModel(input,record,CanEdit,page){Catalogs=await maintenance.CatalogsAsync(Actor,ct)});
    }
    [HttpPost("Firma")]
    [Authorize(Policy="perm:CONFIGURACION.EDITAR")]
    [RequestSizeLimit(2300000)]
    public async Task<IActionResult> UploadSignature(IFormFile? signature,string? rowVersion,CancellationToken ct)
    {
        try
        {
            byte[] version;try{version=Convert.FromBase64String(rowVersion??"");}catch(FormatException){throw new MaintenanceException("Vuelve a abrir la configuración.");}
            if(version.Length!=8)throw new MaintenanceException("Guarda primero los datos de HYT.");
            if(signature is null)throw new MaintenanceException("Selecciona una firma JPG o PNG.");
            await BudgetSignature.ValidateAsync(signature,ct);
            var metadata=await files.SaveAsync(signature,new("firma","FIRMA","Firma","CONFIGURACION"),Guid.NewGuid(),ct,DocumentUploadRules.Folder("FIRMA","Firma"));
            await repository.SaveSignatureAsync(Actor,JsonSerializer.Serialize(metadata),version,ct);
            TempData["Notice"]="Firma guardada en Cloudflare. Se utilizará en las nuevas impresiones de presupuesto.";
        }
        catch(MaintenanceException error){TempData["Notice"]=error.Message;}
        return RedirectToAction(nameof(Index));
    }
    [HttpGet("Firma")]
    public async Task<IActionResult> Signature(CancellationToken ct)
    {
        var file=BudgetSignature.Metadata(await repository.GetAsync(Actor,1,ct));if(file is null)return NotFound();
        try{var bytes=await files.ReadAsync(file,ct);if(bytes is null)return NotFound();Response.Headers.CacheControl="no-store";return File(bytes,file.TipoContenido);}
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpPost("Guardar")]
    [Authorize(Policy="perm:CONFIGURACION.EDITAR")]
    public async Task<IActionResult> Save(MaintenanceEditInput input,CancellationToken ct)
    {
        if(!input.Fields.ContainsKey("SubtituloCotizacion"))input.Fields["SubtituloCotizacion"]=(await repository.GetAsync(Actor,1,ct))?.Text("SubtituloCotizacion")??CompanySettings.DefaultSubtitle;
        var validation=MaintenanceInput.Validate(CompanySettings.Definition,input.Fields,input.Id,input.RowVersion,[]);
        foreach(var issue in validation.Issues)ModelState.AddModelError(issue.Field,issue.Message);
        var ruc=input.Fields.GetValueOrDefault("Ruc")?.Trim();
        if(ruc?.Length!=11||ruc.Any(c=>c is <'0' or >'9'))ModelState.AddModelError("Ruc","El RUC debe tener once dígitos.");
        if(input.Id.HasValue&&input.Id!=1)ModelState.AddModelError("","Configuración no válida.");
        if(ModelState.IsValid)
        {
            try{await repository.SaveAsync(Actor,validation.Payload,validation.Version,ct);TempData["Notice"]="Datos de HYT guardados. Los documentos históricos conservan sus datos.";return RedirectToAction(nameof(Index));}
            catch(MaintenanceException error){ModelState.AddModelError("",error.Message);}
        }
        return View("Index",new CompanyModel(input,await repository.GetAsync(Actor,1,ct),CanEdit,1){Catalogs=await maintenance.CatalogsAsync(Actor,ct)});
    }
}
