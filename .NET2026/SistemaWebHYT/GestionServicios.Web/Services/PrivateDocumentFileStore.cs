using System.Security.Cryptography;
using System.Text.RegularExpressions;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.Services;
public interface IDocumentFileStore
{
    Task<DocumentAttachment> SaveAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct);
    Task<byte[]?> ReadAsync(DocumentAttachment file,CancellationToken ct);
}
public sealed class PrivateDocumentFileStore(IConfiguration configuration,IWebHostEnvironment environment):IDocumentFileStore
{
    public const long MaximumSize=15*1024*1024;
    private string Root
    {
        get
        {
            var configured=configuration["DocumentStorage:Root"];
            var root=Path.GetFullPath(string.IsNullOrWhiteSpace(configured)?Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"SistemaWebHYT","Adjuntos"):configured);
            var web=Path.GetFullPath(environment.WebRootPath??Path.Combine(environment.ContentRootPath,"wwwroot"));
            var comparison=OperatingSystem.IsWindows()?StringComparison.OrdinalIgnoreCase:StringComparison.Ordinal;
            if(root.Equals(web,comparison)||root.StartsWith(web+Path.DirectorySeparatorChar,comparison))throw new MaintenanceException("La carpeta de adjuntos debe estar fuera de la carpeta pública.");
            return root;
        }
    }
    private string Location(string key)
    {
        if(!Regex.IsMatch(key,@"^[a-f0-9]{32}\.(pdf|png|jpg)$",RegexOptions.CultureInvariant))throw new MaintenanceException("La referencia del archivo no es válida.");
        return Path.Combine(Root,key);
    }
    private async Task<DocumentAttachment> SaveCoreAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct)
    {
        var validated=await DocumentUploadRules.ValidateAsync(upload,kind,operation,ct);var bytes=validated.Bytes;var metadata=validated.Metadata;var key=metadata.ClaveAlmacenamiento;var location=Location(key);Directory.CreateDirectory(Root);
        var staging=Path.Combine(Root,"stage-"+Guid.NewGuid().ToString("N")+".tmp");
        try
        {
            await using(var output=new FileStream(staging,FileMode.CreateNew,FileAccess.Write,FileShare.None,81920,FileOptions.Asynchronous))
            {await output.WriteAsync(bytes,ct);await output.FlushAsync(ct);}
            try{File.Move(staging,location,false);}
            catch(IOException)when(File.Exists(location))
            {
                if(new FileInfo(location).Length!=bytes.LongLength)throw new MaintenanceException("La operación ya tiene otro archivo. Vuelve a abrir el registro.");
                var prior=await File.ReadAllBytesAsync(location,ct);
                if(!CryptographicOperations.FixedTimeEquals(SHA256.HashData(prior),SHA256.HashData(bytes)))throw new MaintenanceException("La operación ya tiene otro archivo. Vuelve a abrir el registro para seleccionar uno diferente.");
            }
        }
        finally{if(File.Exists(staging))File.Delete(staging);}
        return metadata;
    }
    private async Task<byte[]?> ReadCoreAsync(DocumentAttachment file,CancellationToken ct)
    {
        var location=Location(file.ClaveAlmacenamiento);if(!File.Exists(location))return null;
        if(file.TamanoBytes is <=0 or >MaximumSize||new FileInfo(location).Length!=file.TamanoBytes)throw new MaintenanceException("El archivo no coincide con su tamaño registrado.");
        var bytes=await File.ReadAllBytesAsync(location,ct);
        if(!Convert.ToHexString(SHA256.HashData(bytes)).Equals(file.SHA256,StringComparison.OrdinalIgnoreCase))throw new MaintenanceException("El archivo no coincide con su huella de integridad.");
        return bytes;
    }
    public async Task<DocumentAttachment> SaveAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct)
    {
        try{return await SaveCoreAsync(upload,kind,operation,ct);}
        catch(Exception error)when(error is IOException or UnauthorizedAccessException){throw new MaintenanceException("No se pudo conservar el adjunto. Revisa el registro antes de reintentar.");}
    }
    public async Task<byte[]?> ReadAsync(DocumentAttachment file,CancellationToken ct)
    {
        try{return await ReadCoreAsync(file,ct);}
        catch(Exception error)when(error is IOException or UnauthorizedAccessException){throw new MaintenanceException("No se pudo recuperar el archivo. Vuelve a intentarlo.");}
    }}