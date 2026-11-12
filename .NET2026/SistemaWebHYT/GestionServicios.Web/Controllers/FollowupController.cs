// Firma: FRANCO LARA - 10/11/2026 | Abre el comprobante emitido vinculado desde la acción Facturado de OC.
// Firma: FRANCO LARA - 06/11/2026 | OC enlazada automáticamente a la revisión de presupuesto asociada, sin selector editable. Listado OC sin columnas de registro, actualización ni responsable/fecha del marcado facturado. Consulta de comprobante y adjuntos OC/confirmación/Guía desde celdas. Estado OC en pop-up con trazabilidad y tarjetas enlazadas a presupuesto/Guía. Confirmación compacta en una sola columna abre consulta con fecha, datos y adjuntos. Nuevas cargas en carpetas de tipo y número de documento. Columna Facturado retirada; borde de confirmación/facturado según registro vigente. Guía compacta Sí/- abre consulta de registro y adjuntos. Guía sin estados, misma ficha en consulta bloqueada y edición. KPI OC y filtros de pendientes; adjuntos de Guía consultables desde grilla. Retira botón duplicado Ver confirmación, conservando consulta en la celda del número. Guía/Ejecución reutilizan ficha de consulta y edición en popup; consulta permite abrir edición.
using System.Security.Claims;
using GestionServicios.Web.Services;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Followup;
using GestionServicios.Web.ViewModels.Shared;
using GestionServicios.Web.ViewModels.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
namespace GestionServicios.Web.Controllers;
// Firma: FRANCO LARA - 03/11/2026 | Listados, consulta, registro y edición documental con permisos, versiones, adjuntos privados y auditoría.
// Firma: FRANCO LARA - 05/11/2026 | Consulta por cliente, tienda y estado con navegación desde el panel.
// Firma: FRANCO LARA - 11/11/2026 | Reportes de OC y guías con los mismos filtros del listado.
[Authorize]
[Route("Seguimiento/{kind}")]
public sealed class FollowupController(IDocumentRepository repository,IDocumentFileRepository fileRepository,IDocumentFileStore files,TimeProvider clock):Controller
{
    private string Actor=>User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DocumentKind? Allowed(string key)=>DocumentKind.All.FirstOrDefault(item=>item.Key==key&&User.HasClaim(AccessService.PermissionClaim,item.Permission+".CONSULTAR"));
    [HttpGet("")]
    public async Task<IActionResult> Index(string kind,string? search,DateOnly? from,DateOnly? to,int page=1,bool? confirmed=null,string? client=null,string? store=null,string? state=null,bool pendingConfirmation=false,bool pendingInvoice=false,CancellationToken ct=default,bool print=false)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();
        if(pendingInvoice&&!User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR"))return Forbid();
        if(definition.Code!="OC"&&(pendingInvoice||pendingConfirmation))return BadRequest("Revisa los filtros del listado.");
        var lima=TimeZoneInfo.FindSystemTimeZoneById("America/Lima");var today=DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),lima).DateTime);
        var first=from??today.AddMonths(-1);var last=to??today;
        if(definition.Code=="GUIA"&&!string.IsNullOrEmpty(state))return BadRequest("La guía no utiliza estados.");
        if(!ModelState.IsValid||client?.Length>200||store?.Length>200||state?.Length>40||page is <1 or >100000000||search?.Length>200||first>last||last==DateOnly.MaxValue)return BadRequest("Revisa los filtros.");
        var start=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(first.ToDateTime(TimeOnly.MinValue),lima));
        var end=new DateTimeOffset(TimeZoneInfo.ConvertTimeToUtc(last.AddDays(1).ToDateTime(TimeOnly.MinValue),lima));
        Task<DocumentPage> Load(int number) => repository.ListAsync(definition,Actor,start,end,search,number,state,
            definition.Code=="OC"?confirmed:null,ct,client,store,pendingConfirmation,pendingInvoice);
        DocumentPage data;
        try
        {
            if (print)
            {
                if (definition.Code is not ("OC" or "GUIA")) return NotFound();
                if (!User.HasClaim(AccessService.PermissionClaim, "REPORTES.IMPRIMIR")) return Forbid();
                var rows = await ListingReport.AllAsync(async number =>
                {
                    var result = await Load(number);
                    return (result.Rows, result.Total);
                }, (System.Text.Json.JsonElement row) => row.GetProperty("Id").GetInt64());
                var summary = new List<string> { ListingReport.Range(first, last), ListingReport.Filter("Buscar", search) };
                if (state is not null) summary.Add(ListingReport.Filter("Estado", ReportDisplay.State(state)));
                if (client is not null) summary.Add(ListingReport.Filter("Cliente", client));
                if (store is not null) summary.Add(ListingReport.Filter("Tienda", store));
                if (definition.Code == "OC" && confirmed is bool confirmation) summary.Add(confirmation ? "Confirmadas" : "Sin confirmar");
                if (pendingConfirmation) summary.Add("Pendiente de confirmación");
                if (pendingInvoice) summary.Add("Pendiente de facturar");
                var table = DocumentReport(definition, rows);
                if (definition.Code == "GUIA")
                    return View("ListingReport", new ListingReport(definition.Title, summary, rows.Count, table));
                var totals = rows.Where(row => ListingReport.Text(row, "CodigoMoneda").Length > 0)
                    .GroupBy(row => ListingReport.Text(row, "CodigoMoneda"))
                    .Select(group => new[] { group.Key, group.Count().ToString(),
                        ListingReport.Money(group.Sum(row => ListingReport.Amount(row, "Importe"))) });
                return View("ListingReport", new ListingReport(definition.Title, summary, rows.Count, table,
                    new ListingTable("Totales por moneda", ["Moneda", "Órdenes", "Importe"], totals)));
            }
            data = await Load(page);
        }
        catch (MaintenanceException error) { return Conflict(error.Message); }
        var states=definition.Code switch{"OC"=>new[]{"REGISTRADA","TRABAJO_COMPLETADO","CERRADA_OPERATIVAMENTE","CANCELADA"},"GUIA"=>["REGISTRADA","ENVIADA","OBSERVADA","ACEPTADA","ANULADA"],"EJECUCION"=>["PENDIENTE","PROGRAMADA","EN_EJECUCION","TERMINADA","CANCELADA"],_=>["VIGENTE","RETIRADA"]};
        var stateFilter=new FilterSelect("state","Estado",state??"",new[]{new FilterChoice("","Todos")}.Concat(states.Select(value=>new FilterChoice(value,value.Replace('_',' ').ToLowerInvariant()))).ToArray());
        var query=new Dictionary<string,string>{["pendingConfirmation"]=pendingConfirmation.ToString().ToLowerInvariant(),["pendingInvoice"]=pendingInvoice.ToString().ToLowerInvariant(),["client"]=client??"",["store"]=store??"",["state"]=state??"",["kind"]=kind,["search"]=search??"",["from"]=first.ToString("yyyy-MM-dd"),["to"]=last.ToString("yyyy-MM-dd"),["confirmed"]=confirmed?.ToString().ToLowerInvariant()??""};
        return View(new DocumentListModel(definition,new GridModel{
            Label=definition.Title,Controller="Followup",Page=page,TotalRecords=data.Total,Query=query,RowQuery=new(){["kind"]=kind},
            Headers=GridHeaders(definition),
            Rows=data.Rows.Select(row=>BuildRow(definition,row)).ToArray()
        },new FilterModel{Texts=definition.Code is "OC" or "GUIA"?[]:[new("client","Cliente (nombre)",client??""),new("store","Tienda (nombre)",store??"")],Checks=definition.Code=="OC"?new[]{new FilterCheck("pendingConfirmation","Pendiente de confirmación del cliente",pendingConfirmation)}.Concat(User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR")?[new FilterCheck("pendingInvoice","Pendiente de facturar",pendingInvoice)]:[]).ToArray():[],Controller="Followup",WithDates=true,From=first,To=last,Search=search??"",Placeholder="Número, presupuesto, cliente o tienda",Selects=definition.Code=="OC"?[stateFilter,new("confirmed","Confirmación del cliente",confirmed?.ToString().ToLowerInvariant()??"",[new("","Todas"),new("true","Confirmadas"),new("false","Sin confirmar")])]:(definition.Code=="GUIA"?[]:[stateFilter])},definition.Code=="OC"?new[]{("REGISTRADA","Registrada"),("TRABAJO_COMPLETADO","Trabajo completado"),("CERRADA_OPERATIVAMENTE","Cerrada operativamente"),("CANCELADA","Cancelada")}.Select(item=>new StateKpi(item.Item1,item.Item2,data.StateCounts?.GetValueOrDefault(item.Item1)??0)).ToArray():[]));
    }
    private ListingTable DocumentReport(DocumentKind kind, IReadOnlyList<System.Text.Json.JsonElement> rows)
    {
        bool Has(string permission) => User.HasClaim(AccessService.PermissionClaim, permission + ".CONSULTAR");
        var guide = Has("GUIAS");
        var order = Has("OC_CLIENTE");
        var confirmation = Has("CONFIRMACIONES");
        var billed = Has("FACTURADO");
        if (kind.Code == "GUIA")
            return new(null, ["Guía / acta", "Fecha de guía", "Presupuesto", "Cliente", "Tienda", "OC del cliente", "Adjuntos"],
                rows.Select(row => new[] { ListingReport.Text(row,"NumeroDocumento"), ListingReport.Date(row,"FechaDocumento"),
                    ListingReport.Text(row,"PresupuestoNumero"), ListingReport.Text(row,"ClienteNombre"), ListingReport.Text(row,"TiendaNombre"),
                    order ? ListingReport.Text(row,"NumeroOCCliente") : "", ListingReport.Text(row,"Adjuntos") }));

        return new(null, ["OC del cliente", "Emisión", "Recepción", "Presupuesto", "Cliente", "Tienda", "Estado", "Moneda", "Importe", "Seguimiento"],
            rows.Select(row =>
            {
                row.TryGetProperty("Facturado", out var invoice);
                var references = new List<string>();
                void Add(string label, string value) { if (value.Length > 0) references.Add(label + " " + value); }
                if (guide) Add("Guía", ListingReport.Text(row,"NumeroGuia"));
                if (confirmation) Add("Confirmación", ListingReport.Text(row,"NumeroConfirmacion"));
                if (billed)
                {
                    var series = ListingReport.Text(invoice,"Serie");
                    var number = ListingReport.Text(invoice,"NumeroDocumento");
                    Add("Comprobante emitido", ListingReport.Text(row,"NumeroComprobante"));
                    if (number.Length > 0) Add("Comprobante", series.Length > 0 ? $"{series}-{number}" : number);
                }
                return new[] { ListingReport.Text(row,"NumeroDocumento"), ListingReport.Date(row,"FechaDocumento"), ListingReport.Date(row,"FechaRecepcionOC"),
                    ListingReport.Text(row,"PresupuestoNumero"), ListingReport.Text(row,"ClienteNombre"), ListingReport.Text(row,"TiendaNombre"),
                    ReportDisplay.State(ListingReport.Text(row,"Estado")), ListingReport.Text(row,"CodigoMoneda"),
                    ListingReport.Money(ListingReport.Amount(row,"Importe")), string.Join("\n", references) };
            }));
    }

    [HttpGet("{id:long}/CambiarEstado")]
    public async Task<IActionResult> StatePopup(string kind,long id,string? returnUrl=null,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();if(id<=0||definition.Code!="OC")return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,id,null,1,ct);if(record is null)return NotFound();
            ViewData["StatePopup"]=true;ViewData["StateReturnUrl"]=Back(definition,returnUrl);
            return PartialView("_StatePopup",new DocumentDetailsModel(definition,record,1));
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpGet("{id:long}/Adjuntos")]
    public async Task<IActionResult> Attachments(string kind,long id,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();if(id<=0)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,id,null,1,ct);
            if(record is null)return NotFound();
            if(definition.Code=="GUIA")return Prepare(definition,record,record.Data.GetProperty("PadreId").GetInt64(),null,true,true);
            return PartialView("_AttachmentPopup",new DocumentDetailsModel(definition,record,1));
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    [HttpGet("{id:long}")]
    public async Task<IActionResult> Details(string kind,long id,int auditPage=1,bool popup=false,string? returnUrl=null,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null)return Forbid();
        if(id<=0||auditPage is <1 or >100000000)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,id,null,auditPage,ct);
            if(record is null)return NotFound();
            if(definition.Code is "GUIA" or "EJECUCION")return Prepare(definition,record,record.Data.GetProperty("PadreId").GetInt64(),returnUrl,popup,true);
            return View(new DocumentDetailsModel(definition,record,auditPage));
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }
    private bool Has(DocumentKind kind, string action) => User.HasClaim(AccessService.PermissionClaim, kind.Permission + "." + action);
    private bool CanEdit(DocumentKind kind) => Has(kind, "MODIFICAR");
    private string Back(DocumentKind kind,string? returnUrl)=>Url.IsLocalUrl(returnUrl)?returnUrl!:Url.Action(nameof(Index),new{kind=kind.Key})!;
    private static byte[]? Version(string? value)
    {
        try{var bytes=Convert.FromBase64String(value??"");return bytes.Length==8?bytes:null;}catch(FormatException){return null;}
    }
    private IActionResult FormView(DocumentFormModel model)=>model.Popup?PartialView("_Form",model):View("Form",model);
    [HttpGet("Registro")]
    public async Task<IActionResult> Register(string kind,long parentId,string? returnUrl,bool popup=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!Has(definition, "CREAR") && !CanEdit(definition))return Forbid();
        if(parentId<=0)return BadRequest();
        try
        {
            var record=await repository.GetAsync(definition,Actor,null,parentId,1,ct);if(record is null)return NotFound();
            var exists = record.Data.TryGetProperty("Documento", out var document) && document.ValueKind == System.Text.Json.JsonValueKind.Object;
            if (!Has(definition, exists ? "MODIFICAR" : "CREAR")) return Forbid();
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
    private IActionResult Prepare(DocumentKind kind,DocumentRecord record,long parentId,string? returnUrl,bool popup,bool readOnly=false)
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
        return FormView(new(kind,input,record,popup,readOnly));
    }
    [HttpPost("Guardar")]
    [RequestSizeLimit(18874368)]
    public async Task<IActionResult> Save(string kind,DocumentInputModel input,bool popup=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!Has(definition, input.Id.HasValue ? "MODIFICAR" : "CREAR"))return Forbid();
        var parentVersion=Version(input.ParentVersion);var originalVersion=Version(input.OriginalVersion);
        if(input.ParentId<=0||input.Id is <=0||input.OperationId==Guid.Empty||parentVersion is null||(input.Id.HasValue&&originalVersion is null))return BadRequest("Vuelve a abrir el registro.");
        if(input.RemovedFiles.Count>200||input.RemovedFiles.Any(id=>id<=0))return BadRequest("Selección de archivos no válida.");
        if(definition.Code=="OC")
        {
            DocumentRecord? context;
            try{context=await repository.GetAsync(definition,Actor,input.Id,input.ParentId,1,ct);}catch(MaintenanceException error){return Conflict(error.Message);}
            if(context is null)return NotFound();
            input.Fields["RevisionId"]=input.Id.HasValue&&context.Data.TryGetProperty("Documento",out var linked)&&linked.ValueKind==System.Text.Json.JsonValueKind.Object
                ?BudgetRecord.Text(linked,"RevisionId"):BudgetRecord.Text(context.Data,"RevisionPresupuestoId");
        }
        var validated=DocumentFields.Validate(definition,input.Fields,input.MotivoCambio,input.Id.HasValue);
        foreach(var issue in validated.Issues)ModelState.AddModelError(issue.Field=="MotivoCambio"?issue.Field:"Fields["+issue.Field+"]",issue.Message);
        input.ReturnUrl=Back(definition,input.ReturnUrl);
        if(ModelState.IsValid)
        {
            try
            {
                if(input.Attachment is not null)
                {
                    var numberField=definition.Code switch{"OC"=>"NumeroOCCliente","GUIA"=>"NumeroActa","CONFIRMACION"=>"NumeroConfirmacion",_=>throw new MaintenanceException("Tipo de adjunto no válido.")};
                    var folder=DocumentUploadRules.Folder(definition.Code,Convert.ToString(validated.Values[numberField])??"");
                    validated.Values["Adjunto"]=await files.SaveAsync(input.Attachment,definition,input.OperationId,ct,folder);
                }
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
        if(kind.Code is "GUIA" or "EJECUCION")actions.Add(new("Ver",Url.Action(nameof(Details),new{kind=kind.Key,id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true));
        if(kind.Code=="OC")actions.Add(new(Has(kind,"CAMBIAR_ESTADO")?"Cambiar estado":"Consultar estado",Url.Action(nameof(StatePopup),new{kind=kind.Key,id,returnUrl=Request.Path+Request.QueryString})!,true));
        if(CanEdit(kind))actions.Add(new("Editar",Url.Action(nameof(Edit),new{kind=kind.Key,id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true));
        var hasConfirmation = long.TryParse(BudgetRecord.Text(row,"ConfirmacionId"),out _);
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.CONSULTAR")&&User.HasClaim(AccessService.PermissionClaim,hasConfirmation?"CONFIRMACIONES.MODIFICAR":"CONFIRMACIONES.CREAR"))
            actions.Add(new("Confirmación del cliente",Url.Action(nameof(Register),new{kind="confirmaciones",parentId=id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true,HasRecord:long.TryParse(BudgetRecord.Text(row,"ConfirmacionId"),out _)));
        if(kind.Code=="OC"&&User.HasClaim(AccessService.PermissionClaim,"FACTURADO.CONSULTAR"))
        {
            if(long.TryParse(BudgetRecord.Text(row,"ComprobanteId"),out var invoiceId))
                actions.Add(new("Facturado",Url.Action("Details","CommercialDocuments",new{scope="ventas",id=invoiceId})!,false,HasRecord:true));
            else
                actions.Add(new("Facturado",Url.Action("Index","InvoiceMark",new{orderId=id,returnUrl=Request.Path+Request.QueryString,popup=true})!,true,HasRecord:row.TryGetProperty("Facturado",out var marked)&&marked.ValueKind==System.Text.Json.JsonValueKind.Object));
        }
        return actions;
    }    [HttpGet("Archivos/{fileId:long}")]
    public async Task<IActionResult> FileDocument(string kind,long fileId,bool download=false,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null || download && !Has(definition, "DESCARGAR"))return Forbid();if(fileId<=0)return BadRequest();
        try
        {
            var file=await fileRepository.GetAsync(definition,Actor,fileId,ct,download);if(file is null)return NotFound();
            var bytes=await files.ReadAsync(file.Metadata,ct);if(bytes is null)return NotFound("El archivo no está disponible.");
            Response.Headers["X-Content-Type-Options"]="nosniff";Response.Headers["Cache-Control"]="private, no-store";
            Response.Headers["Content-Security-Policy"]="sandbox; default-src 'none'; frame-ancestors 'self'";
            return download?File(bytes,file.Metadata.TipoContenido,file.Metadata.NombreOriginal):File(bytes,file.Metadata.TipoContenido);
        }
        catch(MaintenanceException error){return Conflict(error.Message);}
    }    [HttpPost("{id:long}/Estado")]
    public async Task<IActionResult> ChangeState(string kind,long id,long parentId,string? parentVersion,string? originalVersion,Guid operationId,string? state,DateOnly? eventDate,string? responsible,string? reason,bool popup=false,string? returnUrl=null,CancellationToken ct=default)
    {
        var definition=Allowed(kind);if(definition is null||!Has(definition, definition.Code == "OC" ? "CAMBIAR_ESTADO" : "MODIFICAR"))return Forbid();
        if(definition.Code=="GUIA")return BadRequest("La guía no utiliza estados.");
        var parent=Version(parentVersion);var original=Version(originalVersion);
        if(!ModelState.IsValid||id<=0||parentId<=0||parent is null||original is null||operationId==Guid.Empty||eventDate is null||string.IsNullOrWhiteSpace(state)||state.Length>40||string.IsNullOrWhiteSpace(responsible)||responsible.Length>150||reason?.Length>2000)return BadRequest("Revisa la fecha, responsable y estado seleccionado.");
        try{await repository.ChangeStateAsync(definition,Actor,id,parentId,parent,original,operationId,state,eventDate.Value,responsible.Trim(),reason?.Trim(),ct);TempData["DocumentNotice"]="Cambio de estado registrado.";}
        catch(MaintenanceException error){if(popup)return Conflict(error.Message);TempData["DocumentNotice"]=error.Message;}
        if(popup)return Json(new{redirect=Back(definition,returnUrl)});
        return RedirectToAction(nameof(Details),new{kind,id});
    }    private GridRow BuildRow(DocumentKind kind,System.Text.Json.JsonElement row)
    {
        var cellActions=new Dictionary<int,GridCellAction>();
        var cells=new[]{"NumeroInterno","NumeroDocumento","PresupuestoNumero","ClienteNombre","Estado","FechaDocumento","FechaRegistro","UsuarioRegistro","FechaActualizacion","UsuarioActualizacion","Adjuntos"}.Where(name=>(kind.Code!="OC"||name is not ("FechaRegistro" or "UsuarioRegistro" or "FechaActualizacion" or "UsuarioActualizacion"))&&(kind.Code!="GUIA"||name!="Estado")).Select(name=>name=="ClienteNombre"?BudgetRecord.Text(row,name)+" / "+BudgetRecord.Text(row,"TiendaNombre"):DocumentDetailsModel.Display(row,name)).ToList();
        if(kind.Code=="OC")
        {
            if(User.HasClaim(AccessService.PermissionClaim,"PRESUPUESTOS.CONSULTAR")&&long.TryParse(BudgetRecord.Text(row,"PresupuestoId"),out var budgetId))
                cellActions[2]=new(Url.Action("Details","Budget",new{id=budgetId})!,Label:"Abrir presupuesto",Navigate:true);
            cells.Add(DocumentDetailsModel.Display(row,"CodigoMoneda"));cells.Add(DocumentDetailsModel.Display(row,"Importe"));
            var confirmation=BudgetRecord.Text(row,"NumeroConfirmacion");cells.Add(string.IsNullOrEmpty(confirmation)?"Pendiente":confirmation);
            if(long.TryParse(BudgetRecord.Text(row,"Adjuntos"),out var ocFiles)&&ocFiles>0)
                cellActions[6]=new(Url.Action(nameof(Attachments),new{kind="oc-cliente",id=BudgetRecord.Text(row,"Id")})!,Label:"Ver adjuntos de la OC");
            if(User.HasClaim(AccessService.PermissionClaim,"CONFIRMACIONES.CONSULTAR")&&long.TryParse(BudgetRecord.Text(row,"ConfirmacionId"),out var confirmationId))
                cellActions[cells.Count-1]=new(Url.Action(nameof(Attachments),new{kind="confirmaciones",id=confirmationId})!,Label:"Ver confirmación y adjuntos");
            var hasGuide=long.TryParse(BudgetRecord.Text(row,"GuiaId"),out var guideId)&&guideId>0;
            cells.Add(hasGuide?"Sí":"-");
            if(hasGuide&&User.HasClaim(AccessService.PermissionClaim,"GUIAS.CONSULTAR"))
                cellActions[cells.Count-1]=new(Url.Action(nameof(Attachments),new{kind="guias",id=guideId})!,Label:"Ver registro de guía y adjuntos");
        }
        if(kind.Code=="GUIA"&&long.TryParse(BudgetRecord.Text(row,"Adjuntos"),out var filesCount)&&filesCount>0)
            cellActions[cells.Count-1]=new(Url.Action(nameof(Attachments),new{kind="guias",id=BudgetRecord.Text(row,"Id")})!,Label:"Ver adjuntos de la guía");
        return new(BudgetRecord.Text(row,"Id"),cells,CanView:kind.Code is not ("OC" or "GUIA" or "EJECUCION"),Actions:RowActions(kind,row),CellActions:cellActions);
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
        var headers=(kind.Code=="OC"?new[]{"Registro interno","OC del cliente","Presupuesto","Cliente / tienda","Estado","Emisión","Adjuntos","Moneda","Importe","Confirmación del cliente","Guía"}:new[]{"Registro","Documento","Presupuesto","Cliente / tienda","Estado","Fecha","Registro","Registrado por","Actualización","Actualizado por","Adjuntos"}).ToList();
        if(kind.Code=="GUIA")headers.Remove("Estado");
        return headers;
    }}