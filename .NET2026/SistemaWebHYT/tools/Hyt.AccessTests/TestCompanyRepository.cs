using System.Text.Json;
using GestionServicios.Application.Maintenance;
sealed class TestCompanyRepository:ICompanySettingsRepository
{
    private readonly List<BudgetSignatureOption> signatures=[];
    public Task<IReadOnlyList<BudgetSignatureOption>> ListBudgetSignaturesAsync(string actor,CancellationToken ct)
        => Task.FromResult<IReadOnlyList<BudgetSignatureOption>>(signatures);
    public Task SaveBudgetSignatureAsync(string actor,string? name,string? payload,bool defaultSignature,long? id,bool? active,CancellationToken ct)
    {
        if(id is null)
        {
            if(signatures.Count==0)defaultSignature=true;
            signatures.Add(new(signatures.Count+1,name!,payload!,false,true));
            id=signatures.Count;
        }
        for(var i=0;i<signatures.Count;i++)
        {
            var signature=signatures[i];
            if(defaultSignature)signature=signature with{Predeterminada=signature.FirmaId==id};
            if(signature.FirmaId==id&&active.HasValue)signature=signature with{Activo=active.Value,Predeterminada=active.Value&&signature.Predeterminada};
            signatures[i]=signature;
        }
        SignatureJson=signatures.FirstOrDefault(item=>item.Activo&&item.Predeterminada)?.ArchivoJson;
        return Task.CompletedTask;
    }
    private readonly List<DocumentSeries> series=[];
    public Task<IReadOnlyList<DocumentSeries>> ListSeriesAsync(string actor,CancellationToken ct)
        => Task.FromResult<IReadOnlyList<DocumentSeries>>(series);
    public Task AddSeriesAsync(string actor,string type,string value,CancellationToken ct)
    {
        series.Add(new(type,value));
        return Task.CompletedTask;
    }
    public string Subtitle {get;set;}=CompanySettings.DefaultSubtitle;
    public string? SignatureJson {get;set;}
    public Task SaveSignatureAsync(string actor,string payload,byte[] expectedVersion,CancellationToken ct){SignatureJson=payload;return Task.CompletedTask;}
    public int Saves{get;private set;}
    public Task<RecordSnapshot?> GetAsync(string actor,int page,CancellationToken ct)
    {
        var values=JsonSerializer.Deserialize<Dictionary<string,JsonElement>>(JsonSerializer.Serialize(new {Ruc="00000000000",RazonSocial="HYT PRUEBA",Direccion="Dirección de prueba",CodigoUbigeo="150101",Telefono="",Correo="",NombreFirmante="Firmante",CargoFirmante="Comercial",FirmaPresupuestoJson=SignatureJson,SubtituloCotizacion=Subtitle,FechaRegistro="2026-11-03T18:30:00",UsuarioRegistro="Operador",FechaActualizacion=(string?)null,UsuarioActualizacion=(string?)null,VersionFila=Convert.ToBase64String(new byte[8])}))!;
        return Task.FromResult<RecordSnapshot?>(new(1,values,[],[],0));
    }
    public Task SaveAsync(string actor,string payload,byte[]? expectedVersion,CancellationToken ct){Saves++;using var data=JsonDocument.Parse(payload);if(data.RootElement.TryGetProperty("SubtituloCotizacion",out var subtitle))Subtitle=subtitle.GetString()??"";return Task.CompletedTask;}
}
