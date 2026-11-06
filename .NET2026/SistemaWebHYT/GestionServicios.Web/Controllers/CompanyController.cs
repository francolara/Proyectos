// Firma: FRANCO LARA - 06/11/2026 | Catálogos de ubigeo para dirección HYT.
using System.Security.Claims;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 03/11/2026 | Configuración única de HYT con concurrencia y auditoría.
[Authorize(Policy="perm:CONFIGURACION.CONSULTAR")]
[Route("Configuracion/HYT")]
public sealed class CompanyController(ICompanySettingsRepository repository,IMaintenanceRepository maintenance):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit=>User.HasClaim(AccessService.PermissionClaim,"CONFIGURACION.EDITAR");
    [HttpGet("")]
    public async Task<IActionResult> Index(int page=1,CancellationToken ct=default)
    {
        if(page is <1 or >100000000)return BadRequest();var record=await repository.GetAsync(Actor,page,ct);
        var input=new MaintenanceEditInput{Id=record is null?null:1,RowVersion=record?.Text("VersionFila")};
        foreach(var field in CompanySettings.Definition.Fields)input.Fields[field.Name]=record?.Text(field.Name);
        return View(new CompanyModel(input,record,CanEdit,page){Catalogs=await maintenance.CatalogsAsync(Actor,ct)});
    }
    [HttpPost("Guardar")]
    [Authorize(Policy="perm:CONFIGURACION.EDITAR")]
    public async Task<IActionResult> Save(MaintenanceEditInput input,CancellationToken ct)
    {
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
