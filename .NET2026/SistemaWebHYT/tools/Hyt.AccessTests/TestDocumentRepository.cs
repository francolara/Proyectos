using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Followup;
sealed class TestDocumentRepository:IDocumentRepository
{
    public DocumentAttachment? LastFile {get;private set;}
    public bool Retired {get;private set;}
    public int StateChanges {get;private set;}
    public bool RichList {get;set;}
    public bool PendingConfirmation {get;private set;}
    public bool PendingInvoice {get;private set;}
    public int Saves {get;private set;}
    public string? LastPayload {get;private set;}
    public Guid LastOperation {get;private set;}
    public int Page {get;private set;}
    public bool? LastConfirmed {get;private set;}
    public string? Search {get;private set;}
    public DateTimeOffset From {get;private set;}
    private static JsonElement Data(object value)=>JsonSerializer.SerializeToElement(value);
    public Task<DocumentPage> ListAsync(DocumentKind kind,string actor,DateTimeOffset from,DateTimeOffset until,string? search,int page,string? state,bool? confirmed,CancellationToken ct,string? client=null,string? store=null,bool pendingConfirmation=false,bool pendingInvoice=false)
    {
        Page=page;Search=search;From=from;LastConfirmed=confirmed;PendingConfirmation=pendingConfirmation;PendingInvoice=pendingInvoice;
        return Task.FromResult(new DocumentPage([Data(new{Id=1,PresupuestoId=1,NumeroInterno="OC-1",NumeroDocumento="OC-000A/01",PresupuestoNumero="PPTO-1",ClienteNombre="Cliente <script>prueba</script>",TiendaNombre="Tienda",Estado="REGISTRADA",FechaDocumento="2026-11-03",FechaRegistro="2026-11-03T18:30:00Z",UsuarioRegistro="Operador",Adjuntos=RichList?1:0,CodigoMoneda="PEN",Importe=236m,ConfirmacionId=confirmed==true||RichList?(long?)1:null,NumeroConfirmacion=confirmed==true||RichList?"CONF-0001A":null,FechaConfirmacion=confirmed==true||RichList?"2026-11-03":null,AdjuntosConfirmacion=RichList?1:0,GuiaId=RichList?(long?)1:null,NumeroGuia=RichList?"GUIA-0007B":null,AdjuntosGuia=RichList?1:0,Facturado=RichList?new{FechaDocumento="2026-11-03",CodigoTipoDocumento="01",Serie="F001",NumeroDocumento="000012A"}:null})],search == "prueba-reporte" ? 1 : 41,page,new Dictionary<string,long>{{"REGISTRADA",41}}));
    }
    public Task<DocumentRecord?> GetAsync(DocumentKind kind,string actor,long? id,long? parentId,int auditPage,CancellationToken ct)
    {
        if(id==999)return Task.FromResult<DocumentRecord?>(null);
        var document=new Dictionary<string,object?>{[DocumentFields.PrimaryKey(kind)]=id??1,["NumeroRegistroInterno"]="OC-1",["NumeroOCCliente"]="OC-000A/01",["NumeroActa"]="GUIA-0007B",["NumeroConfirmacion"]="ACEP-0002B",["EstadoDocumento"]=kind.Code=="CONFIRMACION"?"VIGENTE":kind.Code=="EJECUCION"?"PENDIENTE":"REGISTRADA",["FechaActa"]="2026-11-03",["FechaEmisionOC"]="2026-11-03",["FechaRecepcionOC"]="2026-11-03",["FechaConfirmacion"]="2026-11-03",["RevisionId"]=1,["FechaRegistro"]="2026-11-03T18:30:00Z",["UsuarioRegistro"]="Operador",["Importe"]=236m,["VersionFila"]=Convert.ToBase64String(new byte[8])};
        var attachments=LastFile is null&&!RichList?Array.Empty<object>():new object[]{new{ArchivoId=100,NombreOriginal=LastFile?.NombreOriginal??"evidencia.pdf",Retirado=Retired,FechaRegistro="2026-11-03T18:30:00Z",UsuarioRegistro="Operador"}};
        return Task.FromResult<DocumentRecord?>(new(Data(new{Adjuntos=attachments,PresupuestoId=1,OrdenCompraId=1,GuiaId=1,EjecucionId=1,ConfirmacionId=1,PadreId=parentId??1,PresupuestoNumero="PPTO-1",ClienteNombre="Cliente",TiendaNombre="Tienda",VersionPadre=Convert.ToBase64String(new byte[8]),RevisionPresupuestoId=1,EstadoPresupuesto="APROBADO",Revisiones=new[]{new{RevisionId=1,NumeroRevision=1,Estado="APROBADO",CodigoMoneda="PEN"}},Documento=parentId==77&&id is null?null:document}),[new BudgetAudit(1,"CREAR",null,"REGISTRADA",new DateTimeOffset(2026,11,3,18,30,0,TimeSpan.Zero),null,"Operador",null,null)],1));
    }
    public Task ChangeStateAsync(DocumentKind kind,string actor,long id,long parentId,byte[] parentVersion,byte[] originalVersion,Guid operation,string state,DateOnly eventDate,string responsible,string? reason,CancellationToken ct)
    {StateChanges++;return Task.CompletedTask;}
    public Task<DocumentSaved> SaveAsync(DocumentKind kind,string actor,long parentId,byte[] parentVersion,string payload,Guid operationId,long? id,byte[]? originalVersion,CancellationToken ct)
    {
        Saves++;LastPayload=payload;LastOperation=operationId;
        using var data=JsonDocument.Parse(payload);
        if(data.RootElement.TryGetProperty("Adjunto",out var attachment)){LastFile=attachment.Deserialize<DocumentAttachment>();Retired=false;}
        if(data.RootElement.TryGetProperty("RetirarAdjuntos",out var retired)&&retired.EnumerateArray().Any(item=>item.GetInt64()==100))Retired=true;
        return Task.FromResult(new DocumentSaved(id??1,LastFile is null?null:100));
    }
}
sealed class TestDocumentFileRepository(TestDocumentRepository documents):IDocumentFileRepository
{
    public Task<DocumentFile?> GetAsync(DocumentKind kind,string actor,long fileId,CancellationToken ct,bool download=false)
        =>Task.FromResult<DocumentFile?>(fileId==100&&kind.Code=="GUIA"&&!documents.Retired&&documents.LastFile is not null?new(100,documents.LastFile):null);
}