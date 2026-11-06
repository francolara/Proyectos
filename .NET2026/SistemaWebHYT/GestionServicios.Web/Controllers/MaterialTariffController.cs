using System.Globalization;
using System.Security.Claims;
using System.Text.RegularExpressions;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Maintenance;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 05/11/2026 | Tarifario simple por material y listado unificado con alta por selección de material, sin versiones ni publicación manual.
[Authorize]
[Route("Materiales/{materialId:long}/Tarifario")]
public sealed class MaterialTariffController(IMaintenanceRepository repository):Controller
{
    [HttpGet("/Tarifarios")]
    public async Task<IActionResult> IndexGlobal(string? search,int page=1,CancellationToken ct=default)
    {
        if(!CanRead)return Forbid();if(page<1||page>100000000||search?.Length>200)return BadRequest();
        ViewData["Search"]=search;ViewData["CanEdit"]=CanEdit;
        return View("Index",await repository.MaterialTariffsAsync(Actor,0,page,ct,search));
    }
    [HttpGet("/Tarifarios/Nuevo")]
    public IActionResult NewGlobal()
    {
        if(!CanEdit)return Forbid();return View("GlobalForm",new MaterialTariffFormModel(0,"",new(),null,false));
    }
    [HttpGet("/Tarifarios/{materialId:long}/{id:long}/Editar")]
    public Task<IActionResult> EditGlobal(long materialId,long id,CancellationToken ct)=>Read(materialId,id,false,ct,true);
    [HttpGet("/Tarifarios/{materialId:long}/{id:long}")]
    public Task<IActionResult> DetailsGlobal(long materialId,long id,CancellationToken ct)=>Read(materialId,id,true,ct,true);
    [HttpPost("/Tarifarios/Guardar")]
    public Task<IActionResult> SaveGlobal(long selectedMaterialId,MaterialTariffInput input,CancellationToken ct)=>Save(selectedMaterialId,input,ct,true);
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanRead=>User.HasClaim(AccessService.PermissionClaim,"MATERIALES.CONSULTAR")&&User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.CONSULTAR");
    private bool CanEdit=>CanRead&&User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.EDITAR");
    private static bool Editable(RecordSnapshot record)=>record.Text("VersionEstado")=="BORRADOR"||record.Text("EsTarifarioMaterial").Equals("true",StringComparison.OrdinalIgnoreCase);
    private Task<RecordSnapshot?> Material(long id,CancellationToken ct)=>repository.GetAsync(MaintenanceDefinitions.Find("materiales")!,Actor,id,cancellationToken:ct);
    [HttpGet("Nuevo")]
    public async Task<IActionResult> New(long materialId,CancellationToken ct)
    {
        if(!CanEdit)return Forbid();if(materialId<=0)return BadRequest();var material=await Material(materialId,ct);if(material is null)return NotFound();
        return View("Form",new MaterialTariffFormModel(materialId,material.Text("Descripcion"),new(),null,false));
    }
    [HttpGet("{id:long}/Editar")]
    public Task<IActionResult> Edit(long materialId,long id,CancellationToken ct)=>Read(materialId,id,false,ct);
    [HttpGet("{id:long}")]
    public Task<IActionResult> Details(long materialId,long id,CancellationToken ct)=>Read(materialId,id,true,ct);
    private async Task<IActionResult> Read(long materialId,long id,bool readOnly,CancellationToken ct,bool standalone=false)
    {
        if(!CanRead||!readOnly&&!CanEdit)return Forbid();if(materialId<=0||id<=0)return BadRequest();
        var material=await Material(materialId,ct);if(material is null)return NotFound();
        var record=await repository.GetAsync(MaintenanceDefinitions.Find("partidas-tarifario")!,Actor,id,cancellationToken:ct);
        if(record is null||record.Text("MaterialId")!=materialId.ToString(CultureInfo.InvariantCulture))return NotFound();
        string Money(string key)=>decimal.TryParse(record.Text(key),NumberStyles.Number,CultureInfo.InvariantCulture,out var value)?value.ToString("0.00",CultureInfo.InvariantCulture):"";
        var input=new MaterialTariffInput{Id=id,RowVersion=record.Text("VersionFila"),TaxCode=record.Text("CodigoTipoIGV"),PricePEN=Money("PrecioPEN"),PriceUSD=Money("PrecioUSD"),Active=record.Text("Activo").Equals("true",StringComparison.OrdinalIgnoreCase)};
        return View(standalone?"GlobalForm":"Form",new MaterialTariffFormModel(materialId,material.Text("Descripcion"),input,record,readOnly||!Editable(record)));
    }
    [HttpPost("Guardar")]
    public async Task<IActionResult> Save(long materialId,MaterialTariffInput input,CancellationToken ct,bool standalone=false)
    {
        if(!CanEdit)return Forbid();if(input.Id is <=0)return BadRequest();
        if(materialId<=0){if(!standalone)return BadRequest();ModelState.AddModelError("","Selecciona un material.");return View("GlobalForm",new MaterialTariffFormModel(0,"",input,null,false));}
        var material=await Material(materialId,ct);if(material is null)return NotFound();
        RecordSnapshot? record=null;byte[]? version=null;
        if(input.Id.HasValue)
        {
            record=await repository.GetAsync(MaintenanceDefinitions.Find("partidas-tarifario")!,Actor,input.Id.Value,cancellationToken:ct);
            if(record is null||record.Text("MaterialId")!=materialId.ToString(CultureInfo.InvariantCulture))return NotFound();
            if(!Editable(record))return BadRequest("Este tarifario conservado no permite edición.");
            try{version=Convert.FromBase64String(input.RowVersion??"");}catch(FormatException){}
            if(version?.Length!=8)ModelState.AddModelError("","Vuelve a abrir el tarifario para actualizarlo.");
        }
        if(input.TaxCode is not ("CON_IGV" or "SIN_IGV"))ModelState.AddModelError(nameof(input.TaxCode),"Selecciona el tipo de IGV.");
        decimal? Money(string? value,string key,bool required)
        {
            value=value?.Trim();if(string.IsNullOrEmpty(value)){if(required)ModelState.AddModelError(key,"Completa el precio en soles.");return null;}
            if(!Regex.IsMatch(value,@"^[0-9]{1,13}(\.[0-9]{1,2})?$")||!decimal.TryParse(value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out var amount))
            {ModelState.AddModelError(key,"Ingresa un precio no negativo, con punto y hasta dos decimales.");return null;}
            return amount;
        }
        var pen=Money(input.PricePEN,nameof(input.PricePEN),false);var usd=Money(input.PriceUSD,nameof(input.PriceUSD),false);
        if(pen is null&&usd is null)ModelState.AddModelError("","Ingresa al menos un precio en soles o dólares.");
        if(ModelState.IsValid)
        {
            try
            {
                await repository.SaveMaterialTariffAsync(Actor,materialId,input.Id,version,input.TaxCode,pen,usd,input.Active,ct);
                if(standalone)return RedirectToAction(nameof(IndexGlobal));
                ViewData["ResetMaterialPricePage"]=!input.Id.HasValue;
                return View("~/Views/Master/CompactSaved.cshtml",materialId);
            }
            catch(MaintenanceException error){ModelState.AddModelError("",error.Message);}
        }
        return View(standalone?"GlobalForm":"Form",new MaterialTariffFormModel(materialId,material.Text("Descripcion"),input,record,false));
    }
}
