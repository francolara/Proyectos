using System.Text.Json;
using GestionServicios.Application.Budgets;

sealed class TestBudgetRepository : IBudgetRepository
{
    public int Saved {get;private set;}
    public int StateChanges {get;private set;}
    public int RequestedPage {get;private set;}
    public DateTimeOffset From {get;private set;}
    public DateTimeOffset Until {get;private set;}
    public string? LastPayload {get;private set;}
    public Task<BudgetPage> ListAsync(string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,CancellationToken ct,string? client=null,string? store=null)
    {
        RequestedPage=page;From=from;Until=until;
        return Task.FromResult(new BudgetPage([new(1,"PPTO-1","Servicio","INVITACION",1,1,1,1,"Cliente <prueba>","Tienda Centro","BORRADOR","PEN",100,18,118,null,DateTimeOffset.UtcNow,"Operador",null,null,null,null,null,null,null,0)],41,page));
    }
    public Task<BudgetRecord?> GetAsync(string actor,long id,long? revisionId,int auditPage,CancellationToken ct)
    {
        if(id==999)return Task.FromResult<BudgetRecord?>(null);
        var revision=new Dictionary<string,object?>();foreach(var field in BudgetDraftRules.HeaderFields)revision[field.Name]=null;
        revision["RevisionId"]=revisionId??id;revision["NumeroRevision"]=1;revision["FechaEmision"]="2026-11-03";revision["CodigoMoneda"]="PEN";revision["CodigoTipoIGV"]="SIN_IGV";revision["Estado"]=id is 2 or 56?"ENVIADO":"BORRADOR";
        revision["ClienteNombre"]="Cliente <prueba>";revision["TiendaNombre"]="Tienda Centro";revision["TasaIGVAplicada"]=0.18m;revision["FechaDesdeTasaIGV"]="2026-11-03";revision["DescripcionTrabajo"]="Servicio de prueba";revision["VersionFila"]=Convert.ToBase64String(new byte[8]);revision["FechaAprobacion"]=null;revision["Subtotal"]=100m;revision["Impuesto"]=18m;revision["Total"]=118m;
        revision["TiendaCodigo"]="T001";revision["SimboloMoneda"]="S/";revision["EmpresaSnapshotJson"]=JsonSerializer.Serialize(new {Ruc="00000000000",RazonSocial="HYT PRUEBA",Direccion="Dirección de prueba",Telefono="",Correo="",NombreFirmante="Firmante",CargoFirmante="Comercial"});
        revision["Consideraciones"]="Consideraciones de prueba.\nIncluye los trabajos descritos.";
        revision["NumeroCotizacionCliente"]="COT-0001A";revision["ClienteDireccion"]="Dirección del cliente";revision["Trimestre"]="IV TRIMESTRE";revision["ZonaTexto"]="Sur / Centro";revision["VigenciaTexto"]="Precios válidos 2026";revision["CodigoSAPProveedor"]="00003000009822";revision["ContratoMarco"]="8000017240";revision["Especialidad"]="INST. SANITARIAS";revision["FechaEjecucion"]="2026-11-10";revision["Coordinador"]="Coordinador";revision["Encargado"]="Encargado";
        var line=new Dictionary<string,object?> { ["PartidaId"]=1,["MaterialId"]=1,["PartidaTarifarioId"]=1,["OrigenPrecio"]="TARIFARIO",["Grupo"]=null,["Posicion"]="001-A",["CodigoInterno"]="0001-A",["CodigoMaterialSAP"]="000SAP",["CentroSAPFICO"]="F001",["NumeroNecesidad"]=null,["TipoServicio"]=null,["Descripcion"]="Servicio",["Unidad"]="UND",["Cantidad"]=1m,["PrecioOrigen"]=100m,["CodigoMonedaOrigen"]="PEN",["CodigoTipoIGVOrigen"]="SIN_IGV",["PrecioUnitario"]=100m,["BaseImponible"]=100m,["Impuesto"]=18m,["Total"]=118m,["ZonasTexto"]="Sur, Centro",["MotivoPrecioManual"]=null };
        line["NumeroLinea"]=1;
        var lines=new[]{line};
        if(id==50)
        {
            lines=Enumerable.Range(1,40).Select(index=>new Dictionary<string,object?>(line){["PartidaId"]=index,["NumeroLinea"]=index,["Grupo"]="Grupo "+((index-1)/10+1),["Descripcion"]="Servicio "+index+". "+string.Concat(Enumerable.Repeat("Descripción extensa de trabajos y materiales; incluye códigos y observaciones técnicas. ",3))}).ToArray();
            revision["Subtotal"]=4000m;revision["Impuesto"]=720m;revision["Total"]=4720m;
        }
        if(id==51){lines=[];revision["Subtotal"]=null;revision["Impuesto"]=null;revision["Total"]=null;}
        if(id==52)
        {revision["CodigoMoneda"]="USD";revision["SimboloMoneda"]="$";revision["CodigoTipoIGV"]="CON_IGV";line["CodigoMonedaOrigen"]="USD";line["CodigoTipoIGVOrigen"]="CON_IGV";line["PrecioOrigen"]=118m;line["PrecioUnitario"]=118m;}
        if(id==54)line["Descripcion"]=string.Concat(Enumerable.Repeat("Detalle técnico de mantenimiento y reposición de componentes. ",22))+"FIN-DESCRIPCION";
        var data=JsonSerializer.SerializeToElement(new { PresupuestoId=id,NumeroInterno="PPTO-"+id,ClienteId=1,TiendaId=1,TipoAtencion=id is 55 or 56?"EMERGENCIA":"INVITACION",Descripcion="Servicio de prueba",FechaSolicitud=(string?)null,FechaRegistro="2026-11-03T18:30:00",UsuarioRegistro="Operador",VersionFila=Convert.ToBase64String(new byte[8]),Revision=revision,Partidas=lines,Revisiones=new[]{new {RevisionId=revisionId??id,NumeroRevision=1,Estado=id is 2 or 56?"ENVIADO":"BORRADOR"}},OrdenCompraId=id is 55 or 56 or 57?(long?)1:null,EstadoOC="REGISTRADA",RevisionOC=id });
        return Task.FromResult<BudgetRecord?>(new(data,[],0));
    }
    public Task<BudgetVersion> SaveAsync(string actor,string payload,BudgetVersion? original,bool newRevision,CancellationToken ct)
    {Saved++;LastPayload=payload;return Task.FromResult(new BudgetVersion(original?.BudgetId??3,3,"PPTO-3",newRevision?2:1,new byte[8],new byte[8]));}
    public Task ChangeStateAsync(string actor,BudgetVersion original,string newState,DateOnly eventDate,string responsible,string? reason,CancellationToken ct)
    {StateChanges++;return Task.CompletedTask;}
    public Task<TariffChoicePage> TariffsAsync(string actor,long clientId,DateOnly date,string currency,string? search,int page,CancellationToken ct,long? materialId=null,long? tariffLineId=null)
        =>Task.FromResult(new TariffChoicePage([new(1,1,"0001-A","001-A","000SAP","Servicio", "UND","TAR1",1,"CON_IGV",currency,currency=="USD"?null:118m,null,null,"Sur, Centro")],21,page));
    public Task<JsonElement> DefaultsAsync(string actor,DateOnly date,long? clientId,long? storeId,CancellationToken ct)
        =>Task.FromResult(JsonSerializer.SerializeToElement(new {Tasa=new {FechaDesde="2026-11-03",Valor=0.18m},Empresa=new {Ruc="00000000000",RazonSocial="HYT PRUEBA",NombreFirmante="Firmante de prueba",CargoFirmante="Comercial",ConsideracionesPresupuesto="Condiciones HYT de prueba"},Cliente=new {ClienteId=clientId,Nombre="Cliente",DireccionFiscal="Av. Cliente 123 / LIMA / LIMA / LIMA"},Tienda=new {TiendaId=storeId,Nombre="Centro",Direccion="Av. Tienda 456 / LIMA / LIMA / LIMA",CentroSAPFICO="F001",ZonaTexto="Sur"},Proveedor=new {CodigoSAPProveedor="000P",ContratoMarco="CM-A"}}));
}
