using System.Text.Json;
using GestionServicios.Application.Billing;

public sealed class TestCommercialDocumentRepository : ICommercialDocumentRepository, IFactoringRepository
{
    public IReadOnlyList<long> LinkedIds { get; private set; } = [];
    public FactoringLot? SavedFactoring { get; private set; }
    public bool WithoutFactoring { get; private set; }
    public bool? HasFactoring { get; private set; }
    public long? LinkedOrderId { get; private set; }
    public CommercialDocument? Saved { get; private set; }
    public Task<JsonElement> ReadAsync(string actor, bool sales, long? id, DateOnly? from, DateOnly? to,
        string? type, string? search, int page, long? clientId, long? budgetId, CancellationToken ct, bool withoutFactoring = false, bool? hasFactoring = null,
        IReadOnlyList<string>? materialCodes = null, bool orderLookup = false, bool invoiceLookup = false)
    {
        WithoutFactoring = withoutFactoring;
        HasFactoring = hasFactoring;
        var doc = new
        {
            ComprobanteId=1L,Tipo="01",Serie="F001",Numero="123",Fecha="2026-11-10",Moneda="PEN",RucHYT="20100000001",
            DocumentoTercero="20600000001",NombreTercero="Tercero de prueba",ClienteId=1L,ProveedorId=1L,FormaPago="Contado",Importe=118m,
            ImporteDetraccion=12m,MonedaDetraccion="PEN",EstadoSunat="PENDIENTE",Importado=true,
            Detalle=new[]{new CommercialLine("A","Producto de prueba","NIU",1,100,100)},
            Pagos=Array.Empty<object>(),Presupuestos=Array.Empty<object>(),Costos=Array.Empty<object>()
        };
        return Task.FromResult(JsonSerializer.SerializeToElement(new
        {
            DetraccionMaterial=materialCodes?.Contains("ITEM-A") == true ? new{Codigo="022",Porcentaje=12m} : null,
            CatalogoDetracciones=new[]{new{Codigo="022",Nombre="Comisión mercantil",Porcentaje=12m}},
            RucHYT="20100000001",TasaIgv=0.18m,Documento=id.HasValue?doc:null,Filas=new[]{new{ComprobanteId=1L,Tipo="01",Serie="F001",Numero="123",Fecha="2026-11-10",NombreTercero="Tercero de prueba",Moneda="PEN",MonedaDetraccion="PEN",Importe=118m,ImporteDetraccion=12m,Detraido=0m,Pagado=0m,EstadoSunat="PENDIENTE",LoteId=hasFactoring == true ? 1L : (long?)null,
                FactoringFechaDesembolso=hasFactoring == true ? "2026-11-12" : null,TotalFilas=21}},
            Cliente=clientId.HasValue?new{ClienteId=1,DocumentoTercero="20600000001",NombreTercero="Tercero de prueba"}:null,
            FacturasNotaCredito=invoiceLookup && clientId == 1 && search is null or "" or "F001" ?
                new[]{new{Referencia="F001-123",Fecha="2026-11-10",Moneda="PEN",Importe=118m}} : [],
            MotivosNotaCredito=new[]{new{Codigo="01",Descripcion="Anulación de la operación"}},
            OpcionesUnidad=new[]{new{Codigo="UND",Descripcion="Unidad",CodigoSunat="NIU"}},
            OpcionesOrdenCompra=orderLookup ? new[]{new{OrdenCompraId=1L,NumeroOCCliente="OC-QA",NumeroPresupuesto="202611000002",NumeroConfirmacion="CONF-QA"}} : [],
            OpcionesCliente=Array.Empty<object>(),OpcionesProveedor=Array.Empty<object>(),
            OpcionesPresupuesto=new[]{new{PresupuestoId=2L,NumeroPresupuesto="202611000002",Descripcion="Trabajo",CodigoMoneda="PEN",Total=100m}},
            OpcionesCosto=new[]{new{CostoProveedorId=2L,NumeroPresupuesto="202611000002",Descripcion="Trabajo uno",Importe=50m},
                new{CostoProveedorId=3L,NumeroPresupuesto="202611000003",Descripcion="Trabajo dos",Importe=40m}}
                .Where(row => string.IsNullOrEmpty(search) || row.NumeroPresupuesto.Contains(search) ||
                    row.Descripcion.Contains(search, StringComparison.OrdinalIgnoreCase)).ToArray(),
            OpcionesCostoPresupuesto=Array.Empty<object>()
        }));
    }
    public Task<JsonElement> ReadFactoringAsync(string actor,long? id,DateOnly? from,DateOnly? to,string? search,int page,bool eligible,CancellationToken ct)
    {
        var invoices = new[]{new{ComprobanteId=1L,Serie="F001",Numero="123",Fecha="2026-11-10",
            FechaVencimiento="2026-12-10",NombreTercero="Cliente de prueba",DocumentoTercero="20600000001",Moneda="PEN",Importe=118m}};
        return Task.FromResult(JsonSerializer.SerializeToElement(new
        {
            Lotes=Array.Empty<object>(), Facturas=invoices, Financieras=new[]{new{Nombre="Financiera de prueba"}},
            Lote=id.HasValue?new{LoteId=id.Value,Financiera="Financiera de prueba",FechaNegociacion="2026-11-10",
                Porcentaje=3.125m,Moneda="PEN",Facturas=invoices}:null
        }));
    }
    public Task<long> SaveFactoringAsync(string actor,FactoringLot lot,CancellationToken ct)
    { SavedFactoring=lot;return Task.FromResult(1L); }

    public Task<long> SaveAsync(string actor,bool sales,CommercialDocument document,long? id,CancellationToken ct)
    { Saved=document;return Task.FromResult(1L); }
    public Task ImportBatchAsync(string actor,bool sales,IReadOnlyList<(string File,CommercialDocument Document)> files,CancellationToken ct)
    { Saved=files.Last().Document;return Task.CompletedTask; }
    public Task LinkAsync(string actor,bool sales,long id,IReadOnlyList<long> ids,long? budgetId,string? description,decimal? amount,CancellationToken ct,Guid? operationId=null)
        { LinkedIds=ids; return Task.CompletedTask; }
    public Task LinkOrderAsync(string actor, long id, long? orderId, CancellationToken ct)
    { LinkedOrderId = orderId; return Task.CompletedTask; }
    public Task<JsonElement> ReadWithholdingsAsync(string actor,string state,string? supplier,string? search,int page,CancellationToken ct)
        => Task.FromResult(JsonSerializer.SerializeToElement(new[]{new{ComprobanteId=1L,Serie="F001",Numero="123",
            Fecha="2026-11-10",NombreTercero="Proveedor de prueba",DocumentoTercero="20600000001",ImporteDetraccion=12m,MonedaDetraccion="PEN",
            Pagado=0m,Saldo=12m,TotalFilas=1}}));
    public Task PayAsync(string actor,long id,bool withholding,DateOnly date,decimal amount,string reference,CancellationToken ct)
        => Task.CompletedTask;
}
