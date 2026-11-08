using System.Text.Json;
using GestionServicios.Application.Billing;
sealed class TestInvoiceMarkRepository:IInvoiceMarkRepository
{
    public bool HasConfirmation {get;set;}=true;
    public int Saves {get;private set;}
    public InvoiceMarkCommand? Last {get;private set;}
    private readonly Dictionary<long,InvoiceMarkCommand> references=[];
    private readonly Dictionary<Guid,long> operations=[];
    private readonly List<(long Order,object Data)> events=[];
    private static object Reference(long id,InvoiceMarkCommand value)=>new{ReferenciaId=id,ConfirmacionId=1,FechaDocumento=value.Date!.Value.ToString("yyyy-MM-dd"),CodigoTipoDocumento=value.Type,Serie=value.Series,NumeroDocumento=value.Number,Vigente=!value.Withdraw,MotivoRetiro=value.Withdraw?value.Reason:null,VersionFila=Convert.ToBase64String(new byte[8]),FechaRegistro="2026-11-03T18:30:00Z",UsuarioRegistro="Operador"};
    public Task<JsonElement?> GetAsync(string actor,long orderId,int page,CancellationToken ct)
    {
        if(orderId==999)return Task.FromResult<JsonElement?>(null);
        var rows=references.Where(pair=>pair.Value.OrderId==orderId).OrderByDescending(pair=>pair.Key).ToArray();
        var active=rows.FirstOrDefault(pair=>!pair.Value.Withdraw);
        return Task.FromResult<JsonElement?>(JsonSerializer.SerializeToElement(new{ConfirmacionId=HasConfirmation?(long?)1:null,OrdenCompraId=orderId,NumeroRegistroInterno="OC-1",NumeroOCCliente="CLI-0001A",VersionFila=Convert.ToBase64String(new byte[8]),Referencia=active.Value is null?null:Reference(active.Key,active.Value),Historial=rows.Skip((page-1)*20).Take(20).Select(pair=>Reference(pair.Key,pair.Value)).ToArray(),TotalHistorial=rows.Length,Auditoria=events.Where(item=>item.Order==orderId).Reverse().Skip((page-1)*20).Take(20).Select(item=>item.Data).ToArray(),TotalAuditoria=events.Count(item=>item.Order==orderId)}));
    }
    public Task<long> SaveAsync(string actor,InvoiceMarkCommand input,CancellationToken ct)
    {
        if(operations.TryGetValue(input.OperationId,out var existing))return Task.FromResult(existing);
        Saves++;Last=input;var id=input.ReferenceId??references.Count+1;var prior=references.GetValueOrDefault(id);references[id]=input;
        events.Add((input.OrderId,new{Accion=input.Withdraw?"RETIRAR_FACTURADO":prior is null?"MARCAR_FACTURADO":"CORREGIR_FACTURADO",Usuario="Operador",FechaRegistro="2026-11-03T18:30:00Z",Motivo=input.Reason,CambiosJson=JsonSerializer.Serialize(new{Anterior=prior?.Number,Nuevo=input.Number})}));
        operations[input.OperationId]=id;return Task.FromResult(id);
    }
}