namespace GestionServicios.Application.Maintenance;

public static class CompanySettings
{
    public static EntityDefinition Definition {get;}=new("empresa-hyt","Datos de HYT","ConfiguracionHYT","CONFIGURACION",
    [new("Ruc","RUC","Text",true,11),new("RazonSocial","Razón social","Text",true,200),new("Direccion","Dirección","Text",false,300),new("CodigoUbigeo","Ubigeo","Ubigeo",false,6),
     new("Telefono","Teléfono","Text",false,50),new("Correo","Correo electrónico","Email",false,256),new("NombreFirmante","Firmante predeterminado","Text",false,150),new("CargoFirmante","Cargo del firmante","Text",false,150),new("ConsideracionesPresupuesto","Consideraciones predeterminadas para presupuestos","LongText",false,20000)],null,false);
}
public interface ICompanySettingsRepository
{
    Task<RecordSnapshot?> GetAsync(string actor,int page,CancellationToken ct);
    Task SaveAsync(string actor,string payload,byte[]? expectedVersion,CancellationToken ct);
}
