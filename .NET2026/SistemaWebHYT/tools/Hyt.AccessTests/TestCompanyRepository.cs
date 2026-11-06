using System.Text.Json;
using GestionServicios.Application.Maintenance;
sealed class TestCompanyRepository:ICompanySettingsRepository
{
    public int Saves{get;private set;}
    public Task<RecordSnapshot?> GetAsync(string actor,int page,CancellationToken ct)
    {
        var values=JsonSerializer.Deserialize<Dictionary<string,JsonElement>>(JsonSerializer.Serialize(new {Ruc="00000000000",RazonSocial="HYT PRUEBA",Direccion="Dirección de prueba",CodigoUbigeo="150101",Telefono="",Correo="",NombreFirmante="Firmante",CargoFirmante="Comercial",FechaRegistro="2026-11-03T18:30:00",UsuarioRegistro="Operador",FechaActualizacion=(string?)null,UsuarioActualizacion=(string?)null,VersionFila=Convert.ToBase64String(new byte[8])}))!;
        return Task.FromResult<RecordSnapshot?>(new(1,values,[],[],0));
    }
    public Task SaveAsync(string actor,string payload,byte[]? expectedVersion,CancellationToken ct){Saves++;return Task.CompletedTask;}
}
