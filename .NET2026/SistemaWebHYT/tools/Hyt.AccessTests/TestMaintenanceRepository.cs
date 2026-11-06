using System.Text.Json;
using GestionServicios.Application.Maintenance;

sealed class TestMaintenanceRepository : IMaintenanceRepository
{
    public int SaveCount { get; private set; }
    public string? LastPayload { get; private set; }
    public int LastPage { get; private set; }
    public string? LastSearch { get; private set; }
    public int Publications { get; private set; }
    public int Copies { get; private set; }
    public bool Unavailable { get; set; }
    public bool ConcurrentSave { get; set; }
    public Task<long> SaveMaterialTariffAsync(string actor,long materialId,long? id,byte[]? version,string tax,decimal? pen,decimal? usd,bool active,CancellationToken ct=default)
    {SaveCount++;LastPayload=JsonSerializer.Serialize(new{MaterialId=materialId,CodigoTipoIGV=tax,PrecioPEN=pen,PrecioUSD=usd,Activo=active});return Task.FromResult(id??3L);}
    public Task<MaterialTariffPage> MaterialTariffsAsync(string actorId,long materialId,int page,CancellationToken ct=default,string? search=null)
        =>Task.FromResult(new MaterialTariffPage([JsonSerializer.SerializeToElement(new{Id=1,MaterialId=1L,MaterialCodigo="MAT-000001",MaterialDescripcion="Material de prueba",VersionId=1,Tarifario="Tarifario de prueba",NumeroVersion=1,Estado="BORRADOR",Editable=true,Posicion="001",CodigoMaterialSAP="SAP-01",CodigoMoneda="PEN",CodigoTipoIGV="SIN_IGV",PrecioPEN=100m,PrecioUSD=(decimal?)null,Activo=true,TipoServicio="Servicio",Zonas="Sur",FechaRegistro="2026-11-05T17:00:00Z",UsuarioRegistro="Operador",FechaActualizacion=(string?)null,UsuarioActualizacion=(string?)null})],1,page));
    public Task<MaintenancePage> ListAsync(EntityDefinition entity, string actorId, string? search, int page, long? parentId = null, bool activeOnly = false, CancellationToken cancellationToken = default)
    {
        LastPage = page; LastSearch = search;
        return Task.FromResult(new MaintenancePage([new(entity.Key == "zonas" ? (page - 1) * 20 + 1 : 1, "0001-A", entity.Key == "zonas" ? page == 1 ? "Sur" : "Centro" : "Cliente <script>prueba</script>", "ACTIVO", "Sur · Centro", DateTimeOffset.UtcNow, "Operador", null, null, entity.Key is "tiendas" or "contactos" ? "CLI-000099" : null, entity.Key is "tiendas" or "contactos" ? "Cliente de prueba" : null)], 41, page));
    }
    public Task<RecordSnapshot?> GetAsync(EntityDefinition entity, string actorId, long id, int auditPage = 1, CancellationToken cancellationToken = default)
    {
        if (Unavailable) throw new TestDatabaseException();
        if (id == 999) return Task.FromResult<RecordSnapshot?>(null);
        var fields = new Dictionary<string, object?> { ["Id"] = id,["MaterialId"]=1,["VersionId"]=1,["TarifarioId"]=1,["CodigoInterno"]="MAT-000001",["Descripcion"]="Material de prueba",["UnidadMedidaId"]=1,["UnidadCodigo"]="UND",["CodigoMoneda"]="PEN",["CodigoTipoIGV"]="SIN_IGV",["PrecioPEN"]=100m,["PrecioUSD"]=null, ["Codigo"] = "C1", ["TipoEntidad"] = "JURIDICA", ["RazonSocial"] = "Cliente", ["Activo"] = true,
            ["VersionFila"] = Convert.ToBase64String(new byte[8]), ["Estado"] = id == 2 ? "VIGENTE" : "BORRADOR", ["VersionEstado"] = id == 2 ? "VIGENTE" : "BORRADOR", ["FechaRegistro"] = "2026-11-03T18:30:00", ["UsuarioRegistro"] = "Operador",
            ["CodigoMonedaPredeterminada"] = "PEN", ["CodigoTipoIGVPredeterminado"] = "SIN_IGV", ["Posicion"] = "001-A", ["CodigoMaterialSAP"] = "000SAP", ["Zonas"] = new[] { new { Id = 1, Descripcion = "Sur" }, new { Id = 2, Descripcion = "Centro" } } };
        if(entity.Key=="clientes")fields["Enlaces"]=new[]{new{Id=41L,Descripcion="Portal de documentos",Url="https://example.com/documentos"},new{Id=42L,Descripcion="Carpeta del cliente",Url="https://example.com/carpeta"}};
        var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(JsonSerializer.Serialize(fields))!;
        return Task.FromResult<RecordSnapshot?>(new(id, values, [1, 2], [new("CREAR", "Operador", DateTimeOffset.UtcNow, null, "ACTIVO", "{\"dato\":\"<script>\"}")], 21));
    }
    public Task<long> SaveAsync(EntityDefinition entity, string actorId, long? id, byte[]? version, string payload, CancellationToken cancellationToken = default)
    { if (ConcurrentSave) throw new MaintenanceException("El registro cambió. Vuelve a consultarlo."); SaveCount++; LastPayload = payload; return Task.FromResult(id ?? 3); }
    public Task<IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>>> CatalogsAsync(string actorId, CancellationToken cancellationToken = default)
        => Task.FromResult<IReadOnlyDictionary<string, IReadOnlyList<CatalogOption>>>(new Dictionary<string, IReadOnlyList<CatalogOption>>
        { ["Departments"]=[new("15","LIMA"),new("07","CALLAO")],["Provinces"]=[new("1501","LIMA"),new("0701","CALLAO")],["Districts"]=[new("150101","LIMA"),new("150132","SAN JUAN DE LURIGANCHO"),new("070101","CALLAO")], ["EntityType"] = [new("JURIDICA", "Persona jurídica"), new("NATURAL", "Persona natural")], ["Currency"] = [new("PEN", "Soles")], ["TaxTreatment"] = [new("SIN_IGV", "Sin IGV")] });
    public Task PublishVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default) { Publications++; return Task.CompletedTask; }
    public Task<long> CloneVersionAsync(string actorId, long versionId, byte[] expectedVersion, CancellationToken cancellationToken = default) { Copies++; return Task.FromResult(3L); }
}

sealed class TestDatabaseException : System.Data.Common.DbException
{
    public TestDatabaseException() : base("Datos internos de prueba que no deben aparecer en la vista") { }
}
