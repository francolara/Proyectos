// Firma: FRANCO LARA - 06/11/2026 | Encabezados Dirección y Cliente simplificados.
using System.Security.Claims;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Security;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Maintenance;
using GestionServicios.Web.ViewModels.Shared;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 05/11/2026 | Códigos internos no editables, abreviaturas y tarifas contextuales y datos HYT redirigidos y columnas de cliente y enlaces múltiples y acceso de tarifario simplificado por material y navegación al listado unificado. Dirección completa con ubigeo en los listados de clientes y tiendas. Predeterminados de material con unidad y códigos.
[Authorize]
[Route("Maestros/{entity}")]
public sealed class MasterController(IMaintenanceRepository repository, ILogger<MasterController> logger) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private bool Allowed(EntityDefinition definition, string action) => User.HasClaim(AccessService.PermissionClaim, definition.Permission + "." + action);
    private string Back(string? url, string entity) => Url.IsLocalUrl(url) ? url! : Url.Action(nameof(Index), new { entity })!;
    private static bool Frozen(EntityDefinition entity, RecordSnapshot record) =>
        entity.Key == "versiones-tarifario" && record.Text("Estado") != "BORRADOR" ||
        entity.Key == "partidas-tarifario" && record.Text("VersionEstado") != "BORRADOR";
    private static bool ValidPage(int page) => page is >= 1 and <= 100000000;

    [HttpGet("")]
    public async Task<IActionResult> Index(string entity, string? search, int page = 1, long? parentId = null, CancellationToken ct = default)
    {
        if(entity=="tarifarios")return RedirectToAction("IndexGlobal","MaterialTariff",new{search,page});
        if(entity=="proveedor-cliente")return parentId is >0
            ? RedirectToAction(nameof(Details),new{entity="clientes",id=parentId.Value})
            : RedirectToAction(nameof(Index),new{entity="clientes"});
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null) return NotFound();
        if (!Allowed(definition, "CONSULTAR")) return Forbid();
        if (!ValidPage(page) || search?.Length > 200 || parentId <= 0) return BadRequest();
        MaintenancePage data = new([], 0, page);
        string? error = null;
        try { data = await repository.ListAsync(definition, Actor, search, page, parentId, cancellationToken: ct); }
        catch (Exception ex) when (ex is not OperationCanceledException)
        { logger.LogError(ex, "No se pudo consultar {Entity}", entity); error = "No se pudo cargar el listado. Vuelve a intentarlo."; }
        var query = new Dictionary<string, string> { ["entity"] = entity, ["search"] = search ?? "" };
        if (parentId.HasValue) query["parentId"] = parentId.Value.ToString();
        var returnUrl = Url.Action(nameof(Index), new { entity, search, page, parentId })!;
        var canEdit = Allowed(definition, "EDITAR");
        var showClient = entity is "tiendas" or "contactos";
        return View(new MaintenanceListModel
        {
            Entity = definition, Search = search ?? "", ParentId = parentId, CanEdit = canEdit,
            Grid = new GridModel
            {
                Label = definition.Title, Controller = "Master", Query = query, RowQuery = new() { ["entity"] = entity, ["returnUrl"] = returnUrl },
                Headers = (entity=="unidades"?new[]{"Código interno","Abreviatura"}:new[]{entity=="versiones-tarifario"?"Versión":entity=="partidas-tarifario"?"Posición":"Código interno"}).Concat(new[]{"Descripción"}).Concat(showClient?new[]{"Código de cliente",entity=="tiendas"?"Cliente":"Cliente: nombres / razón social"}:Array.Empty<string>()).Concat(new[]{"Estado",entity is "clientes" or "tiendas"?"Dirección":"Información","Registro","Registrado por","Actualización","Actualizado por"}).ToArray(),
                Rows = data.Rows.Select(row => new GridRow(row.Id.ToString(), (entity=="unidades"?new[]{row.Id.ToString()}:Array.Empty<string>()).Concat(new[]{row.Code, row.Description}).Concat(showClient?new[]{row.ClientCode??"—",row.ClientName??"—"}:Array.Empty<string>()).Concat(new[]{row.Status, row.Extra ?? "", PeruDate.Display(row.CreatedAt), row.CreatedBy,
                    row.UpdatedAt.HasValue ? PeruDate.Display(row.UpdatedAt.Value) : "—", row.UpdatedBy ?? "—"}).ToArray(), true, canEdit)).ToArray(),
                TotalRecords = data.Total, Page = page, Error = error
            }
        });
    }

    [HttpGet("Nuevo")]
    public async Task<IActionResult> New(string entity, long? parentId, string? returnUrl, bool compact=false,long? materialId=null,CancellationToken ct=default)
    {
        if(entity=="tarifarios"&&!compact)return RedirectToAction("NewGlobal","MaterialTariff");
        if(compact&&materialId.HasValue&&entity is "partidas-tarifario" or "tarifarios" or "versiones-tarifario")return RedirectToAction("New","MaterialTariff",new{materialId});
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null) return NotFound();
        if (!Allowed(definition, "EDITAR") || !Allowed(definition, "CONSULTAR")) return Forbid();
        if (parentId <= 0) return BadRequest();
        var input = new MaintenanceEditInput { Compact=compact,MaterialContextId=materialId,ReturnUrl = Back(returnUrl, entity) };
        foreach (var field in definition.Fields) input.Fields[field.Name] = field.Default;
        if (definition.ParentField is not null && parentId.HasValue) input.Fields[definition.ParentField] = parentId.Value.ToString();
        if (entity == "partidas-tarifario" && parentId.HasValue)
        {
            var parent = await repository.GetAsync(MaintenanceDefinitions.Find("versiones-tarifario")!, Actor, parentId.Value, cancellationToken: ct);
            if (parent is null) return NotFound();
            if (parent.Text("Estado") != "BORRADOR") return BadRequest("La versión publicada no permite agregar partidas.");
            input.Fields["CodigoMoneda"] = parent.Text("CodigoMonedaPredeterminada");
            input.Fields["CodigoTipoIGV"] = parent.Text("CodigoTipoIGVPredeterminado");
            input.ZoneIds = parent.Zones.ToList();
        }
        return await Form(definition, input, null, false, 1, ct);
    }

    [HttpGet("{id:long}/Editar")]
    public Task<IActionResult> Edit(string entity, long id, string? returnUrl,bool compact=false,long? materialId=null,CancellationToken ct=default) => Read(entity, id, false, returnUrl, 1, ct,compact,materialId);
    [HttpGet("{id:long}")]
    public Task<IActionResult> Details(string entity, long id, string? returnUrl, int auditPage = 1,bool compact=false,long? materialId=null,CancellationToken ct = default) => Read(entity, id, true, returnUrl, auditPage, ct,compact,materialId);
    private async Task<IActionResult> Read(string entity, long id, bool readOnly, string? returnUrl, int auditPage, CancellationToken ct,bool compact=false,long? materialId=null)
    {
        if(compact&&materialId.HasValue&&entity=="partidas-tarifario")return RedirectToAction(readOnly?"Details":"Edit","MaterialTariff",new{materialId,id});
        if(compact&&materialId.HasValue&&entity is "tarifarios" or "versiones-tarifario")return RedirectToAction("New","MaterialTariff",new{materialId});
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null) return NotFound();
        if (!Allowed(definition, "CONSULTAR") || !readOnly && !Allowed(definition, "EDITAR")) return Forbid();
        if (id <= 0 || !ValidPage(auditPage)) return BadRequest();
        var record = await repository.GetAsync(definition, Actor, id, auditPage, ct);
        if (record is null) return NotFound();
        var input = new MaintenanceEditInput { Compact=compact,MaterialContextId=materialId,Id = id, ReturnUrl = Back(returnUrl, entity), RowVersion = record.Text("VersionFila"), ZoneIds = record.Zones.ToList() };
        foreach (var field in definition.Fields) input.Fields[field.Name] = record.Text(field.Name);
        if(entity=="clientes" && record.Values.TryGetValue("Enlaces",out var links) && links.ValueKind==System.Text.Json.JsonValueKind.Array)
            foreach(var link in links.EnumerateArray())input.DocumentLinks.Add(new(){Id=link.GetProperty("Id").GetInt64(),Description=link.GetProperty("Descripcion").GetString(),Url=link.GetProperty("Url").GetString()});
        return await Form(definition, input, record, readOnly || Frozen(definition, record), auditPage, ct);
    }

    private async Task<IActionResult> Form(EntityDefinition entity, MaintenanceEditInput input, RecordSnapshot? record, bool readOnly, int auditPage, CancellationToken ct, string? error = null)
    {
        RecordSnapshot? material=null;
        if(input.MaterialContextId is long context)
        {
            if(context<=0||entity.Key is not ("tarifarios" or "versiones-tarifario" or "partidas-tarifario"))return BadRequest();
            if(!User.HasClaim(AccessService.PermissionClaim,"MATERIALES.CONSULTAR"))return Forbid();
            material=await repository.GetAsync(MaintenanceDefinitions.Find("materiales")!,Actor,context,cancellationToken:ct);
            if(material is null)return NotFound();
            if(entity.Key=="partidas-tarifario")
            {
                if(record is not null&&record.Text("MaterialId")!=context.ToString())return NotFound();
                input.Fields["MaterialId"]=context.ToString();
                if(record is null&&HttpMethods.IsGet(Request.Method)&&string.IsNullOrEmpty(input.Fields.GetValueOrDefault("CodigoMaterialSAP")))input.Fields["CodigoMaterialSAP"]=material.Text("CodigoMaterialSAP");
            }
        }
        var catalogs = await repository.CatalogsAsync(Actor, ct);
        var prices=entity.Key=="materiales"&&record is not null&&User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.CONSULTAR")
            ?new MaterialTariffsModel(record.Id,record.Text("Descripcion"),await repository.MaterialTariffsAsync(Actor,record.Id,1,ct),User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.EDITAR")):null;
        return View(input.Compact?"CompactForm":"Form", new MaintenanceFormModel { Entity = entity, Input = input, Record = record, Catalogs = catalogs,
            MaterialContextLabel=material is null?null:material.Text("CodigoInterno")+" · "+material.Text("Descripcion"),MaterialTariffs=prices,ReadOnly = readOnly, CanEdit = Allowed(entity, "EDITAR") && (record is null || !Frozen(entity, record)), AuditPage = auditPage, Error = error });
    }

    [HttpPost("Guardar")]
    public async Task<IActionResult> Save(string entity, MaintenanceEditInput input, CancellationToken ct)
    {
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null) return NotFound();
        if (!Allowed(definition, "EDITAR") || !Allowed(definition, "CONSULTAR")) return Forbid();
        input.ReturnUrl = Back(input.ReturnUrl, entity);
        RecordSnapshot? record = null;
        if (input.Id.HasValue)
        {
            record = await repository.GetAsync(definition, Actor, input.Id.Value, cancellationToken: ct);
            if (record is null) return NotFound();
            if (Frozen(definition, record)) return BadRequest("La versión publicada no permite cambios.");
        }
        if(input.MaterialContextId is long context)
        {
            if(!input.Compact||context<=0||entity is not ("tarifarios" or "versiones-tarifario" or "partidas-tarifario"))return BadRequest();
            if(!User.HasClaim(AccessService.PermissionClaim,"MATERIALES.CONSULTAR"))return Forbid();
            if(await repository.GetAsync(MaintenanceDefinitions.Find("materiales")!,Actor,context,cancellationToken:ct) is null)return NotFound();
            if(entity=="partidas-tarifario"&&(input.Fields.GetValueOrDefault("MaterialId")!=context.ToString()||record is not null&&record.Text("MaterialId")!=context.ToString()))return BadRequest("El precio debe corresponder al material de esta ficha.");
        }
        var validated = MaintenanceInput.Validate(definition, input.Fields, input.Id, input.RowVersion, input.ZoneIds, input.DocumentLinks);
        foreach (var issue in validated.Issues) ModelState.AddModelError(issue.Field.Length == 0 ? "" : issue.Field.StartsWith("DocumentLinks",StringComparison.Ordinal) ? issue.Field : "Fields[" + issue.Field + "]", issue.Message);
        if (ModelState.IsValid)
        {
            try
            {
                var id = await repository.SaveAsync(definition, Actor, input.Id, validated.Version, validated.Payload, ct);
                if(input.Compact)
                {
                    if(input.MaterialContextId.HasValue&&entity is "tarifarios" or "versiones-tarifario")
                        return RedirectToAction(nameof(New),new{entity=entity=="tarifarios"?"versiones-tarifario":"partidas-tarifario",parentId=id,compact=true,materialId=input.MaterialContextId,returnUrl=input.ReturnUrl});
                    ViewData["ResetMaterialPricePage"]=!input.Id.HasValue;
                    return View("CompactSaved",input.MaterialContextId??0L);
                }
                TempData["Notice"] = "Registro guardado.";
                return RedirectToAction(nameof(Details), new { entity, id, returnUrl = input.ReturnUrl });
            }
            catch (MaintenanceException ex) { ModelState.AddModelError("", ex.Message); }
        }
        return await Form(definition, input, record, false, 1, ct);
    }

    [HttpGet("{id:long}/Tarifas")]
    public async Task<IActionResult> MaterialTariffs(string entity,long id,int page=1,CancellationToken ct=default)
    {
        if(entity!="materiales"||id<=0||!ValidPage(page))return BadRequest();
        if(!User.HasClaim(AccessService.PermissionClaim,"MATERIALES.CONSULTAR")||!User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.CONSULTAR"))return Forbid();
        var material=await repository.GetAsync(MaintenanceDefinitions.Find("materiales")!,Actor,id,cancellationToken:ct);
        if(material is null)return NotFound();
        return PartialView("_MaterialTariffs",new MaterialTariffsModel(id,material.Text("Descripcion"),await repository.MaterialTariffsAsync(Actor,id,page,ct),User.HasClaim(AccessService.PermissionClaim,"TARIFARIOS.EDITAR")));
    }
    [HttpGet("Ayuda")]
    public async Task<IActionResult> Lookup(string entity, string? search, int page = 1, long? parentId = null, CancellationToken ct = default)
    {
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null) return NotFound();
        if (!Allowed(definition, "CONSULTAR")) return Forbid();
        if (!ValidPage(page) || search?.Length > 200 || parentId <= 0) return BadRequest();
        var data = await repository.ListAsync(definition, Actor, search, page, parentId, true, ct);
        return Json(new { rows = data.Rows.Select(row => new { row.Id, row.Code, row.Description, row.Extra }), total = data.Total, page });
    }

    [HttpGet("{id:long}/Predeterminados")]
    public async Task<IActionResult> Defaults(string entity, long id, CancellationToken ct)
    {
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null || entity is not ("versiones-tarifario" or "materiales")) return NotFound();
        if (!Allowed(definition, "CONSULTAR")) return Forbid();
        if (id <= 0) return BadRequest();
        var record = await repository.GetAsync(definition, Actor, id, cancellationToken: ct);
        if (record is null) return NotFound();
        if (entity == "versiones-tarifario" && record.Text("Estado") != "BORRADOR") return BadRequest();
        var fields = entity == "materiales" ? new[] { "CodigoMaterialSAP", "CodigoInterno", "Descripcion", "UnidadCodigo" } : new[] { "CodigoMonedaPredeterminada", "CodigoTipoIGVPredeterminado" };
        var zones = record.Values.TryGetValue("Zonas", out var data) ? data : System.Text.Json.JsonSerializer.SerializeToElement(record.Zones.Select(zone => new { Id = zone }));
        return Json(new { fields = fields.ToDictionary(field => field, record.Text), zones });
    }

    [HttpPost("{id:long}/Publicar")]
    public Task<IActionResult> Publish(string entity, long id, string? rowVersion, bool confirmed, string? returnUrl,bool compact=false,long? materialId=null,CancellationToken ct=default)
        => ChangeVersion(entity, id, rowVersion, confirmed, returnUrl, false, ct,compact,materialId);

    [HttpPost("{id:long}/Copiar")]
    public Task<IActionResult> Clone(string entity, long id, string? rowVersion, bool confirmed, string? returnUrl,bool compact=false,long? materialId=null,CancellationToken ct=default)
        => ChangeVersion(entity, id, rowVersion, confirmed, returnUrl, true, ct,compact,materialId);

    private async Task<IActionResult> ChangeVersion(string entity, long id, string? rowVersion, bool confirmed, string? returnUrl, bool clone, CancellationToken ct,bool compact=false,long? materialId=null)
    {
        var definition = MaintenanceDefinitions.Find(entity);
        if (definition is null || entity != "versiones-tarifario") return NotFound();
        if (!Allowed(definition, "EDITAR") || !Allowed(definition, "CONSULTAR")) return Forbid();
        byte[]? version = null;
        try { version = Convert.FromBase64String(rowVersion ?? ""); } catch (FormatException) { }
        if (id <= 0 || version?.Length != 8 || !confirmed) return BadRequest("Confirma la operación desde la ficha de la versión.");
        try
        {
            var destination = id;
            if (clone) destination = await repository.CloneVersionAsync(Actor, id, version, ct);
            else await repository.PublishVersionAsync(Actor, id, version, ct);
            if(compact&&materialId.HasValue){ViewData["ResetMaterialPricePage"]=clone;return View("CompactSaved",materialId.Value);}
            TempData["Notice"] = clone ? "Se creó una nueva versión borrador con copia de sus partidas." : "Versión publicada. Sus precios quedan protegidos.";
            return RedirectToAction(nameof(Details), new { entity, id = destination, returnUrl = Back(returnUrl, entity) });
        }
        catch (MaintenanceException ex)
        {
            TempData["Notice"] = ex.Message;
            return RedirectToAction(nameof(Details), new { entity, id,compact,materialId,returnUrl = Back(returnUrl, entity) });
        }
    }
}
