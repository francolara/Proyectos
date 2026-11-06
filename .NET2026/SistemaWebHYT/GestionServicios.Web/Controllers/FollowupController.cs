using System.Security.Claims;
using GestionServicios.Web.Services;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Followup;
using GestionServicios.Web.ViewModels.Shared;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Listados, consulta, registro y edición documental con permisos, versiones, adjuntos privados y auditoría.
// Firma: FRANCO LARA - 05/11/2026 | Consulta por cliente, tienda y estado con navegación desde el panel.
[Authorize]
[Route("Seguimiento/{kind}")]
public sealed class FollowupController(IDocumentRepository repository,IDocumentFileRepository fileRepository,IDocumentFileStore files,TimeProvider clock):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DocumentKind? Allowed(string key)=>DocumentKind.All.FirstOrDefault(item=>item.Key==key&&User.HasClaim(AccessService.PermissionClaim,item.Permission+".CONSULTAR"));
    [HttpGet("")]
    public async Task<IActionResult> Index(string kind,string? search,DateOnly? from,DateOnly? to,int page=1,bool? confirmed=null,string? client=null,string? store=null,string? state=null,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();
        var lima=TimeZoneInfo.FindSystemTimeZoneById("America/Lima");var today=DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),lima).DateTime);
        var first=from??today.AddMonths(-1);var last=to??today;
        if(!ModelState.IsValid||client?.Length>200||store?.Length>200||state?.Length>40||page is <1 or >100000000||search?.Length>200||first>last||last==DateOnly.MaxValue)return BadRequest("Revisa los filtros.");
        var start=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(first.ToDateTime(TimeOnly.MinValue),lima));
        var end=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(last.AddDays(1).ToDateTime(TimeOnly.MinValue),lima));
        DocumentPage data;try{data=await repository.ListAsync(definition,Actor,start,end,search,page,state,definition.Code=="OC"?confirmed:null,ct,client,store);}catch(MaintenanceException error){return Conflict(error.Message);}
        var states=definition.Code switch{"OC"=>new[]{"REGISTRADA","TRABAJO_COMPLETADO","CERRADA_OPERATIVAMENTE","CANCELADA"},"GUIA"=>["REGISTRADA","ENVIADA","OBSERVADA","ACEPTADA","ANULADA"],"EJECUCION"=>["PENDIENTE","PROGRAMADA","EN_EJECUCION","TERMINADA","CANCELADA"],_=>["VIGENTE","RETIRADA"]};
        var stateFilter=new FilterSelect("state","Estado",state??"",new[]{new FilterChoice("","Todos")}.Concat(states.Select(value=>new FilterChoice(value,value.Replace('_',' ').ToLowerInvariant()))).ToArray());
        var query=new Dictionary<string,string>{["client"]=client??"",["store"]=store??"",["state"]=state??"",["kind"]=kind,["search"]=search??"",["from"]=first.ToString("yyyy-MM-dd"),["to"]=last.ToString("yyyy-MM-dd"),["confirmed"]=confirmed?.ToString().ToLowerInvariant()??""};
        return View(new DocumentListModel(definition,new GridModel{
            Label=definition.Title,Controller="Followup",Page=page,TotalRecords=data.Total,Query=query,RowQuery=new(){["kind"]=kind},
            Headers=GridHeaders(definition),
            Rows=data.Rows.Select(row=>BuildRow(definition,row)).ToArray()
        },new FilterModel{Texts=[new("client","Cliente (nombre)",client??""),new("store","Tienda (nombre)",store??"")],Controller="Followup",WithDates=true,From=first,To=last,Search=search??"",Placeholder="Número, presupuesto, cliente o tienda",Selects=definition.Code=="OC"?[stateFilter,new("confirmed","Confirmación del cliente",confirmed?.ToString().ToLowerInvariant()??"",[new("","Todas"),new("true","Confirmadas"),new("false","Sin confirmar")])]:[stateFilter]}));
    }
    [HttpGet("{id:long}")]
    public async Task<IActionResult> Details(string kind,long id,int auditPage=1,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();
        if(id<=0||auditPage is <1 or >100000000)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,id,null,auditPage,ct);
            return record is null?NotFound():View(new DocumentDetailsModel(definition,record,auditPage));
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    private bool CanEdit(DocumentKind kind)=>User.HasClaim(AccessService.PermissionClaim,kind.Permission+".EDITAR");
    private string Back(DocumentKind kind,string? returnUrl)=>Url.IsLocalUrl(returnUrl)?returnUrl!:Url.Action(nameof(Index),new{kind=kind.Key})!;
    private static byte[]? Version(string? value)
    {
        try{var bytes=Convert.FromBase64String(value??"");return bytes.Length==8?bytes:null;}catch(FormatException){return null;}
    }
    private IActionResult FormView(DocumentFormModel model)=>model.Popup?PartialView("_Form",model):View("Form",model);
    [HttpGet("Registro")]
    public async Task<IActionResult> Register(string kind,long parentId,string? returnUrl,bool popup=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!CanEdit(definition))return Forbid();
        if(parentId<=0)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,null,parentId,1,ct);if(record is null)return NotFound();
            return Prepare(definition,record,parentId,returnUrl,popup);
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpGet("{id:long}/Editar")]
    public async Task<IActionResult> Edit(string kind,long id,string? returnUrl,bool popup=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!CanEdit(definition))return Forbid();if(id<=0)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,id,null,1,ct);if(record is null)return NotFound();
            return Prepare(definition,record,record.Data.GetProperty("PadreId").GetInt64(),returnUrl,popup);
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    private IActionResult Prepare(DocumentKind kind,DocumentRecord record,long parentId,string? returnUrl,bool popup)
    {
        var input=new DocumentInputModel{OperationId=Guid.NewGuid(),ParentId=parentId,ParentVersion=BudgetRecord.Text(record.Data,"VersionPadre"),ReturnUrl=Back(kind,returnUrl)};
        if(record.Data.TryGetProperty("Documento",out var document)&&document.ValueKind==System.Text.Json.JsonValueKind.Object)
        {
            input.Id=document.GetProperty(DocumentFields.PrimaryKey(kind)).GetInt64();input.OriginalVersion=BudgetRecord.Text(document,"VersionFila");
            foreach(var field in DocumentFields.For(kind))
            {
                var value=BudgetRecord.Text(document,field.Name);
                input.Fields[field.Name]=field.Kind=="Date"&&value.Length>=10?value[..10]:value;
            }
        }
        else
        {
            var today=DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime).ToString("yyyy-MM-dd");
            foreach(var field in DocumentFields.For(kind))input.Fields[field.Name]=field.Required&&field.Kind=="Date"?today:null;
            if(kind.Code=="OC")input.Fields["RevisionId"]=BudgetRecord.Text(record.Data,"RevisionPresupuestoId");
        }
        return FormView(new(kind,input,record,popup));
    }
    [HttpPost("Guardar")]
    [RequestSizeLimit(18874368)]
    public async Task<IActionResult> Save(string kind,DocumentInputModel input,bool popup=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!CanEdit(definition))return Forbid();
        var parentVersion=Version(input.ParentVersion);var originalVersion=Version(input.OriginalVersion);
        if(input.ParentId<=0||input.Id is <=0||input.OperationId==Guid.Empty||parentVersion is null||(input.Id.HasValue&&originalVersion is null))return BadRequest("Vuelve a abrir el registro.");
        if(input.RemovedFiles.Count>200||input.RemovedFiles.Any(id=>id<=0))return BadRequest("Selección de archivos no válida.");
        var validated=DocumentFields.Validate(definition,input.Fields,input.MotivoCambio,input.Id.HasValue);
        foreach(var issue in validated.Issues)ModelState.AddModelError(issue.Field=="MotivoCambio"?issue.Field:"Fields["+issue.Field+"]",issue.Message);
        input.ReturnUrl=Back(definition,input.ReturnUrl);
        if(ModelState.IsValid)
        {
            try
            {
                if(input.Attachment is not null)validated.Values["Adjunto"]=await files.SaveAsync(input.Attachment,definition,input.OperationId,ct);
                validated.Values["RetirarAdjuntos"]=input.RemovedFiles.Distinct().ToArray();
                var saved=await repository.SaveAsync(definition,Actor,input.ParentId,parentVersion,System.Text.Json.JsonSerializer.Serialize(validated.Values),input.OperationId,input.Id,originalVersion,ct);
                if(popup)return Json(new{redirect=input.ReturnUrl});
                return RedirectToAction(nameof(Details),new{kind,id=saved.Id});
            }
            catch(MaintenanceException error){ModelState.AddModelError("",error.Message);}
        }
        DocumentRecord? record;
        try{record=await repository.GetAsync(definition,Actor,input.Id,input.ParentId,1,ct);}catch(MaintenanceException error){return Conflict(error.Message);}
        if(record is null)return NotFound();
        // Mantener las versiones enviadas: nunca reemplazar un token obsoleto por uno actual al reintentar.
        Response.StatusCode=422;return FormView(new(definition,input,record,popup));
    }    private IReadOnlyList<GridAction> RowActions(DocumentKind kind,System.Text.Json.JsonElement row)
    {
        var id=BudgetRecord.Text(row,"Id");var actions=new List<GridAction>();
        if(CanEdit(kind))actions.Add(new("Editar",Url.Action(nameof(Edit),new{kind=kind.Key,id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true));
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.CONSULTAR")&&User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.EDITAR"))
            actions.Add(new("Confirmación del cliente",Url.Action(nameof(Register),new{kind="confirmaciones",parentId=id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true));
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.CONSULTAR")&&long.TryParse(BudgetRecord.Text(row,"ConfirmacionId"),out var confirmationId))
            actions.Add(new("Ver confirmación",Url.Action(nameof(Details),new{kind="confirmaciones",id=confirmationId})!));
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR"))
            actions.Add(new("Facturado",Url.Action("Index","InvoiceMark",new{orderId=id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true));
        return actions;
    }    [HttpGet("Archivos/{fileId:long}")]
    public async Task<IActionResult> FileDocument(string kind,long fileId,bool download=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();if(fileId<=0)return BadRequest();
        try
        {
            var file=await fileRepository.GetAsync(definition,Actor,fileId,ct);if(file is null)return NotFound();
            var bytes=await files.ReadAsync(file.Metadata,ct);if(bytes is null)return NotFound("El archivo no está disponible.");
            Response.Headers["X-Content-Type-Options"]="nosniff";Response.Headers["Cache-Control"]="private, no-store";
            Response.Headers["Content-Security-Policy"]="sandbox; default-src 'none'; frame-ancestors 'self'";
            return download?File(bytes,file.Metadata.TipoContenido,file.Metadata.NombreOriginal):File(bytes,file.Metadata.TipoContenido);
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }    [HttpPost("{id:long}/Estado")]
    public async Task<IActionResult> ChangeState(string kind,long id,long parentId,string? parentVersion,string? originalVersion,Guid operationId,string? state,DateOnly? eventDate,string? responsible,string? reason,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!CanEdit(definition))return Forbid();
        var parent=Version(parentVersion);var original=Version(originalVersion);
        if(!ModelState.IsValid||id<=0||parentId<=0||parent is null||original is null||operationId==Guid.Empty||eventDate is null||string.IsNullOrWhiteSpace(state)||state.Length>40||string.IsNullOrWhiteSpace(responsible)||responsible.Length>150||reason?.Length>2000)return BadRequest("Revisa la fecha, responsable y estado seleccionado.");
        try{await repository.ChangeStateAsync(definition,Actor,id,parentId,parent,original,operationId,state,eventDate.Value,responsible.Trim(),reason?.Trim(),ct);TempData["DocumentNotice"]="Cambio de estado registrado.";}
        catch(MaintenanceException error){TempData["DocumentNotice"]=error.Message;}
        return RedirectToAction(nameof(Details),new{kind,id});
    }    private GridRow BuildRow(DocumentKind kind,System.Text.Json.JsonElement row)
    {
        var cells=new[]{"NumeroInterno","NumeroDocumento","PresupuestoNumero","ClienteNombre","Estado","FechaDocumento","FechaRegistro","UsuarioRegistro","FechaActualizacion","UsuarioActualizacion","Adjuntos"}.Select(name=>name=="ClienteNombre"?BudgetRecord.Text(row,name)+" / "+BudgetRecord.Text(row,"TiendaNombre"):DocumentDetailsModel.Display(row,name)).ToList();
        if(kind.Code=="OC")
        {
            cells.Add(DocumentDetailsModel.Display(row,"CodigoMoneda"));cells.Add(DocumentDetailsModel.Display(row,"Importe"));
            var confirmation=BudgetRecord.Text(row,"NumeroConfirmacion");cells.Add(string.IsNullOrEmpty(confirmation)?"Pendiente":confirmation);
            cells.Add(DocumentDetailsModel.Display(row,"FechaConfirmacion"));cells.Add(DocumentDetailsModel.Display(row,"AdjuntosConfirmacion"));
        }
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR"))
        {
            var exists=row.TryGetProperty("Facturado",out var invoice)&&invoice.ValueKind==System.Text.Json.JsonValueKind.Object;
            cells.Add(exists?"Sí":"No");
            foreach(var name in new[]{"FechaDocumento","CodigoTipoDocumento","Serie","NumeroDocumento","FechaRegistro","UsuarioRegistro"})cells.Add(exists?DocumentDetailsModel.Display(invoice,name):"—");
        }
        return new(BudgetRecord.Text(row,"Id"),cells,Actions:RowActions(kind,row));
    }    [HttpGet("Expediente")]
    public async Task<IActionResult> Context(string kind,long parentId,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();if(parentId<=0)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,null,parentId,1,ct);
            if(record is null||!record.Data.TryGetProperty("Documento",out var document)||document.ValueKind!=System.Text.Json.JsonValueKind.Object)return NotFound("Todavía no se ha registrado este documento.");
            return RedirectToAction(nameof(Details),new{kind,id=document.GetProperty(DocumentFields.PrimaryKey(definition)).GetInt64()});
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }    private IReadOnlyList<string> GridHeaders(DocumentKind kind)
    {
        var headers=(kind.Code=="OC"?new[]{"Registro interno","OC del cliente","Presupuesto","Cliente / tienda","Estado","Emisión","Registro","Registrado por","Actualización","Actualizado por","Adjuntos","Moneda","Importe","Confirmación del cliente","Fecha confirmación","Adjuntos de confirmación"}:new[]{"Registro","Documento","Presupuesto","Cliente / tienda","Estado","Fecha","Registro","Registrado por","Actualización","Actualizado por","Adjuntos"}).ToList();
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR"))headers.AddRange(["Facturado","Fecha comprobante","Tipo","Serie","Número","Marcado el","Marcado por"]);
        return headers;
    }}