// Firma: FRANCO LARA - 06/11/2026 | Consulta de tarifas genéricas con cliente pendiente; cambio de moneda automático en detalle neto.
using System.Globalization;
using System.Security.Claims;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Budgets;
using GestionServicios.Web.ViewModels.Shared;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 03/11/2026 | Presupuestos MVC con filtros UTC, ayudas, validación y acciones auditadas.
// Firma: FRANCO LARA - 05/11/2026 | Consulta por cliente, tienda y estado con navegación desde el panel. Direcciones completas y consideraciones predeterminadas al crear presupuesto. Detalle neto editable y ayuda de tarifas por material.
[Authorize(Policy="perm:PRESUPUESTOS.CONSULTAR")]
[Route("Presupuestos")]
public sealed class BudgetController(IBudgetRepository repository,TimeProvider clock):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool CanEdit=>User.HasClaim(AccessService.PermissionClaim,"PRESUPUESTOS.EDITAR");
    private static readonly TimeZoneInfo Lima=TimeZoneInfo.FindSystemTimeZoneById("America/Lima");
    private DateOnly Today=>DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),Lima).DateTime);
    private string Back(string? url)=>Url.IsLocalUrl(url)?url!:Url.Action(nameof(Index))!;
    private static bool Date(string? text,out DateOnly date)=>DateOnly.TryParseExact(text,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out date);
    private static string Money(decimal? amount)=>amount?.ToString("N2",CultureInfo.GetCultureInfo("es-PE"))??"Precio pendiente";

    [HttpGet("")]
    public async Task<IActionResult> Index(string? search,DateOnly? from,DateOnly? to,int page=1,string? client=null,string? store=null,string? state=null,CancellationToken ct=default)
    {
        var first=from??Today.AddMonths(-1);var last=to??Today;
        if(!ModelState.IsValid||(state is not (null or "BORRADOR" or "ENVIADO" or "OBSERVADO" or "APROBADO" or "RECHAZADO" or "VENCIDO" or "CANCELADO"))||client?.Length>200||store?.Length>200||state?.Length>40||page is <1 or >100000000||search?.Length>200||first>last||last==DateOnly.MaxValue)return BadRequest("Revisa los filtros.");
        var start=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(first.ToDateTime(TimeOnly.MinValue),Lima));
        var end=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(last.AddDays(1).ToDateTime(TimeOnly.MinValue),Lima));
        var data=await repository.ListAsync(Actor,start,end,search,page,state,ct,client,store);
        var returnUrl=Url.Action(nameof(Index),new {search,client,store,state,from=first.ToString("yyyy-MM-dd"),to=last.ToString("yyyy-MM-dd"),page})!;
        return View(new BudgetListModel(new GridModel
        {
            Label="Presupuestos",Controller="Budget",Page=page,TotalRecords=data.Total,
            Headers=["Presupuesto","Revisión","Cliente / tienda","Tipo","Estado","Moneda","Total","OC cliente","Guia (Acta de conformidad)","Registro","Registrado por","Actualización","Actualizado por","Aprobación"],
            Rows=data.Rows.Select(row=>new GridRow(row.BudgetId.ToString(),[row.Number,row.RevisionNumber.ToString(),row.ClientName+" / "+row.StoreName,row.Type,row.State,row.Currency,Money(row.Total),row.CustomerOrderNumber??"Sin OC",row.GuideNumber??"Sin guía",
                PeruDate.Display(row.CreatedAt),row.CreatedBy,row.UpdatedAt.HasValue?PeruDate.Display(row.UpdatedAt.Value):"—",row.UpdatedBy??"—",row.ApprovedOn?.ToString("dd/MM/yyyy")??"—"],true,CanEdit&&row.State=="BORRADOR"&&!row.PurchaseOrderId.HasValue,DocumentActions(row,returnUrl))).ToArray(),
            Query=new() { ["client"]=client??"",["store"]=store??"",["state"]=state??"",["search"]=search??"",["from"]=first.ToString("yyyy-MM-dd"),["to"]=last.ToString("yyyy-MM-dd") },RowQuery=new() { ["returnUrl"]=returnUrl }
        },new FilterModel { Texts=[new("client","Cliente (nombre)",client??""),new("store","Tienda (nombre)",store??"")],Selects=[new("state","Estado",state??"",new[]{new FilterChoice("","Todos"),new("BORRADOR","Borrador"),new("ENVIADO","Enviado"),new("OBSERVADO","Observado"),new("APROBADO","Aprobado"),new("RECHAZADO","Rechazado"),new("VENCIDO","Vencido"),new("CANCELADO","Cancelado")})],Controller="Budget",Search=search??"",WithDates=true,From=first,To=last,Placeholder="Número, descripción, cliente o tienda" },CanEdit));
    }

    [HttpGet("Nuevo")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    public async Task<IActionResult> New(string? returnUrl,CancellationToken ct)
    {
        var input=new BudgetInputModel {ReturnUrl=Back(returnUrl)};
        foreach(var field in BudgetDraftRules.HeaderFields)input.Draft.Revision[field.Name]=null;
        input.Draft.Revision["FechaEmision"]=Today.ToString("yyyy-MM-dd");input.Draft.Revision["CodigoMoneda"]="PEN";input.Draft.Revision["CodigoTipoIGV"]="SIN_IGV";
        var defaults=await repository.DefaultsAsync(Actor,Today,null,null,ct);
        if(defaults.TryGetProperty("Empresa",out var company)&&company.ValueKind==JsonValueKind.Object)
        { input.Draft.Revision["Consideraciones"]=BudgetRecord.Text(company,"ConsideracionesPresupuesto"); }
        return View("Form",new BudgetFormModel(input,null,defaults,false,CanEdit,"Sin seleccionar","Sin seleccionar"));
    }

    [HttpGet("{id:long}")]
    public Task<IActionResult> Details(long id,long? revisionId,string? returnUrl,int auditPage=1,CancellationToken ct=default)=>Read(id,revisionId,returnUrl,true,false,auditPage,ct);
    [HttpGet("{id:long}/Editar")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    public Task<IActionResult> Edit(long id,long? revisionId,string? returnUrl,CancellationToken ct)=>Read(id,revisionId,returnUrl,false,false,1,ct);
    [HttpGet("{id:long}/NuevaRevision")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    public Task<IActionResult> NewRevision(long id,long revisionId,string? returnUrl,CancellationToken ct)=>Read(id,revisionId,returnUrl,false,true,1,ct);
    private async Task<IActionResult> Read(long id,long? revisionId,string? returnUrl,bool readOnly,bool newRevision,int auditPage,CancellationToken ct)
    {
        if(id<=0||revisionId<=0||auditPage is <1 or >100000000)return BadRequest();
        var record=await repository.GetAsync(Actor,id,revisionId,auditPage,ct);if(record is null)return NotFound();
        var revision=record.Revision;var state=BudgetRecord.Text(revision,"Estado");
        var latest=record.Data.GetProperty("Revisiones").EnumerateArray().First().GetProperty("RevisionId").GetInt64()==revision.GetProperty("RevisionId").GetInt64();
        var hasOrder=record.Data.GetProperty("OrdenCompraId").ValueKind!=JsonValueKind.Null;
        if(!readOnly&&(!latest||hasOrder||newRevision&&(state is "BORRADOR" or "CANCELADO")||!newRevision&&state!="BORRADOR"))return BadRequest("Abre la última revisión disponible para esta acción.");
        var regularize=hasOrder&&record.Text("TipoAtencion")=="EMERGENCIA"&&record.Text("EstadoOC")!="CANCELADA"&&record.Text("RevisionOC")==BudgetRecord.Text(revision,"RevisionId")&&state is "BORRADOR" or "ENVIADO";
        var input=FromRecord(record);input.ReturnUrl=Back(returnUrl);input.NewRevision=newRevision;
        var defaults=await repository.DefaultsAsync(Actor,Date(input.Draft.Revision["FechaEmision"],out var date)?date:Today,input.Draft.ClienteId,input.Draft.TiendaId,ct);
        return View("Form",new BudgetFormModel(input,record,defaults,readOnly,CanEdit&&latest&&!hasOrder,BudgetRecord.Text(revision,"ClienteNombre"),BudgetRecord.Text(revision,"TiendaNombre"),auditPage,CanEdit&&latest&&(!hasOrder||regularize)));
    }
    [NonAction]
    public static BudgetInputModel FromRecord(BudgetRecord record)
    {
        var revision=record.Revision;
        var input=new BudgetInputModel {Id=record.Data.GetProperty("PresupuestoId").GetInt64(),RevisionId=revision.GetProperty("RevisionId").GetInt64(),HeaderVersion=record.Text("VersionFila"),RevisionVersion=BudgetRecord.Text(revision,"VersionFila")};
        input.Draft.ClienteId=record.Data.GetProperty("ClienteId").GetInt64();input.Draft.TiendaId=record.Data.GetProperty("TiendaId").GetInt64();
        input.Draft.TipoAtencion=record.Text("TipoAtencion");input.Draft.Descripcion=BudgetRecord.Text(revision,"DescripcionTrabajo");
        if(string.IsNullOrEmpty(input.Draft.Descripcion))input.Draft.Descripcion=record.Text("Descripcion");
        input.Draft.FechaSolicitud=record.Text("FechaSolicitud") is {Length:>=10} requested?requested[..10]:null;
        foreach(var field in BudgetDraftRules.HeaderFields)input.Draft.Revision[field.Name]=BudgetRecord.Text(revision,field.Name);
        foreach(var line in record.Data.GetProperty("Partidas").EnumerateArray())
        {
            long? Id(string name)=>line.GetProperty(name).ValueKind==JsonValueKind.Null?null:line.GetProperty(name).GetInt64();
            string Text(string name)=>BudgetRecord.Text(line,name);
            input.Draft.Partidas.Add(new() {PartidaId=Id("PartidaId"),PartidaTarifarioId=Id("PartidaTarifarioId"),MaterialId=Id("MaterialId"),OrigenPrecio=Text("OrigenPrecio"),Grupo=Text("Grupo"),Posicion=Text("Posicion"),CodigoInterno=Text("CodigoInterno"),CodigoMaterialSAP=Text("CodigoMaterialSAP"),CentroSAPFICO=Text("CentroSAPFICO"),NumeroNecesidad=Text("NumeroNecesidad"),TipoServicio=Text("TipoServicio"),Descripcion=Text("Descripcion"),Unidad=Text("Unidad"),
                Cantidad=line.GetProperty("Cantidad").GetDecimal().ToString("0.00",CultureInfo.InvariantCulture),PrecioManual=line.GetProperty("PrecioOrigen").ValueKind==JsonValueKind.Null?null:line.GetProperty("PrecioOrigen").GetDecimal().ToString("0.00",CultureInfo.InvariantCulture),CodigoMonedaOrigen=Text("CodigoMonedaOrigen"),CodigoTipoIGVOrigen=Text("CodigoTipoIGVOrigen"),MotivoPrecioManual=Text("MotivoPrecioManual"),ZonasTexto=Text("ZonasTexto")});
        }
        var appliedRate=revision.GetProperty("TasaIGVAplicada").GetDecimal();
        foreach(var line in input.Draft.Partidas)
        {
            if(decimal.TryParse(line.PrecioManual,NumberStyles.Number,CultureInfo.InvariantCulture,out var price)&&line.CodigoTipoIGVOrigen=="CON_IGV")line.PrecioManual=BudgetAmounts.Round(price/(1+appliedRate)).ToString("0.00",CultureInfo.InvariantCulture);
            line.CodigoTipoIGVOrigen="SIN_IGV";line.OrigenPrecio="MANUAL";line.MotivoPrecioManual="Precio neto del detalle del presupuesto";
        }
        input.Draft.Revision["CodigoTipoIGV"]="SIN_IGV";
        return input;
    }

    [HttpPost("Guardar")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    [RequestFormLimits(ValueCountLimit=20000)]
    public async Task<IActionResult> Save(BudgetInputModel input,CancellationToken ct)
    {
        if(input.Draft.DetalleNeto)input.Draft.RevalorizarMoneda=true;
        input.ReturnUrl=Back(input.ReturnUrl);BudgetVersion? original=null;
        try {original=input.Original();}catch(FormatException){ModelState.AddModelError("","Vuelve a abrir el registro; falta su versión original.");}
        foreach(var issue in BudgetDraftRules.Validate(input.Draft))ModelState.AddModelError("Draft."+issue.Field,issue.Message);
        if(input.NewRevision&&original is null)ModelState.AddModelError("","Selecciona la revisión de origen.");
        if(ModelState.IsValid)
        {
            try
            {
                var saved=await repository.SaveAsync(Actor,BudgetDraftRules.Payload(input.Draft),original,input.NewRevision,ct);
                TempData["Notice"]="Presupuesto guardado. Los importes se calcularon en el servidor.";
                return RedirectToAction(nameof(Details),new {id=saved.BudgetId,revisionId=saved.RevisionId,returnUrl=input.ReturnUrl});
            }
            catch(MaintenanceException error){ModelState.AddModelError("",error.Message);}
        }
        var defaults=await repository.DefaultsAsync(Actor,Date(input.Draft.Revision.GetValueOrDefault("FechaEmision"),out var date)?date:Today,input.Draft.ClienteId,null,ct);
        return View("Form",new BudgetFormModel(input,null,defaults,false,CanEdit,"Cliente "+input.Draft.ClienteId,"Tienda "+input.Draft.TiendaId));
    }

    [HttpGet("Tarifario")]
    public async Task<IActionResult> Tariffs(long clientId,string? date,string currency,string? search,int page=1,CancellationToken ct=default,long? materialId=null,long? tariffLineId=null)
    {
        if(materialId<=0||tariffLineId<=0||clientId<0||!Date(date,out var emission)||currency is not ("PEN" or "USD")||search?.Length>200||page is <1 or >100000000)return BadRequest();
        return Json(await repository.TariffsAsync(Actor,clientId,emission,currency,search,page,ct,materialId,tariffLineId));
    }
    [HttpPost("Calcular")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    [RequestFormLimits(ValueCountLimit=20000)]
    public async Task<IActionResult> Calculate(BudgetInputModel input,CancellationToken ct)
    {
        if(input.Draft.Partidas.Count>500||!Date(input.Draft.Revision.GetValueOrDefault("FechaEmision"),out var date))return BadRequest();
        var defaults=await repository.DefaultsAsync(Actor,date,input.Draft.ClienteId,null,ct);
        decimal? rate=defaults.TryGetProperty("Tasa",out var tax)&&tax.ValueKind==JsonValueKind.Object?tax.GetProperty("Valor").GetDecimal():null;
        if(!input.NewRevision&&input.Id is >0&&input.RevisionId is >0)
        {
            var record=await repository.GetAsync(Actor,input.Id.Value,input.RevisionId,1,ct);
            if(record is null)return NotFound();
            if(BudgetRecord.Text(record.Revision,"FechaEmision")==date.ToString("yyyy-MM-dd"))rate=record.Revision.GetProperty("TasaIGVAplicada").GetDecimal();
        }
        if(rate is null)return BadRequest("No hay tasa IGV para la fecha elegida.");
        var lines=new List<BudgetLineAmounts>();
        try
        {
            foreach(var line in input.Draft.Partidas)
            {
                if(!BudgetDraftRules.Number(line.Cantidad,out var quantity))return BadRequest();
                decimal? price=null;
                if(!string.IsNullOrEmpty(line.PrecioManual))
                { if(!BudgetDraftRules.Number(line.PrecioManual,out var parsed))return BadRequest("Revisa los precios: hasta dos decimales.");price=parsed; }
                lines.Add(BudgetAmounts.Calculate(quantity,price,line.CodigoMonedaOrigen,line.CodigoTipoIGVOrigen,input.Draft.Revision.GetValueOrDefault("CodigoMoneda")??"",input.Draft.Revision.GetValueOrDefault("CodigoTipoIGV")??"",rate.Value));
            }
            return Json(new {lines,totals=BudgetAmounts.Sum(lines),rate});
        }
        catch(ArgumentException error){return BadRequest(error.Message);}
    }
    [HttpGet("DatosIniciales")]
    public async Task<IActionResult> Defaults(string? date,long? clientId,long? storeId,CancellationToken ct)
    {
        if(!Date(date,out var emission)||clientId<=0||storeId<=0)return BadRequest();
        return Json(await repository.DefaultsAsync(Actor,emission,clientId,storeId,ct));
    }
    [HttpPost("Estado")]
    [Authorize(Policy="perm:PRESUPUESTOS.EDITAR")]
    public async Task<IActionResult> State(BudgetStateInput input,CancellationToken ct)
    {
        if(input.NewState is not ("ENVIADO" or "OBSERVADO" or "APROBADO" or "RECHAZADO" or "VENCIDO" or "CANCELADO"))return BadRequest("Acción no válida.");
        if(!Date(input.EventDate,out var eventDate)||string.IsNullOrWhiteSpace(input.Responsible)||input.Responsible.Length>150||input.Reason?.Length>2000)return BadRequest("Completa fecha y responsable; revisa la observación.");
        BudgetVersion? original;
        try {original=new BudgetInputModel {Id=input.Id,RevisionId=input.RevisionId,HeaderVersion=input.HeaderVersion,RevisionVersion=input.RevisionVersion}.Original();}
        catch(FormatException){return BadRequest("Vuelve a abrir la revisión.");}
        if(original is null)return BadRequest();
        try {await repository.ChangeStateAsync(Actor,original,input.NewState,eventDate,input.Responsible,input.Reason,ct);TempData["Notice"]="Cambio de estado registrado.";}
        catch(MaintenanceException error){TempData["Notice"]=error.Message;}
        return RedirectToAction(nameof(Details),new {id=input.Id,revisionId=input.RevisionId,returnUrl=Back(input.ReturnUrl)});
    }
    private IReadOnlyList<GridAction> DocumentActions(BudgetListItem row,string returnUrl)
    {
        var actions=new List<GridAction>();
        foreach(var (kind,permission,label) in new[]{("guias","GUIAS","Guía / acta"),("oc-cliente","OC_CLIENTE","OC del cliente"),("ejecucion","EJECUCION","Ejecución")})
        {
            if(!User.HasClaim(AccessService.PermissionClaim,permission+".CONSULTAR"))continue;
            if(User.HasClaim(AccessService.PermissionClaim,permission+".EDITAR"))
                actions.Add(new(label,Url.Action("Register","Followup",new{kind,parentId=row.BudgetId,returnUrl,popup=true})!,true));
            else actions.Add(new(label,Url.Action("Context","Followup",new{kind,parentId=row.BudgetId})!));
        }
        return actions;
    }}
