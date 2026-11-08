// Firma: FRANCO LARA - 06/11/2026 | Rutas por tipo/número con segmentos seguros y lectura compatible de claves históricas.
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.Services;
public sealed record ValidatedDocumentUpload(DocumentAttachment Metadata,byte[] Bytes);
public static class DocumentUploadRules
{
    public const long MaximumSize=15*1024*1024;
    public static async Task<ValidatedDocumentUpload> ValidateAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct,string? folder=null)
    {
        if(kind.Code=="EJECUCION"||operation==Guid.Empty)throw new MaintenanceException("Este registro no admite adjuntos.");
        if(upload.Length is <=0 or >MaximumSize)throw new MaintenanceException("El archivo debe tener contenido y no superar 15 MB.");
        var name=Path.GetFileName(upload.FileName.Replace('\\','/')).Trim();
        if(string.IsNullOrWhiteSpace(name)||name.Length>255||name.Any(char.IsControl))throw new MaintenanceException("Revisa el nombre del archivo; admite hasta 255 caracteres.");
        await using var input=upload.OpenReadStream();using var memory=new MemoryStream();var buffer=new byte[81920];int read;
        while((read=await input.ReadAsync(buffer,ct))>0){if(memory.Length+read>MaximumSize)throw new MaintenanceException("El archivo supera 15 MB.");await memory.WriteAsync(buffer.AsMemory(0,read),ct);}
        var bytes=memory.ToArray();
        if(bytes.LongLength!=upload.Length)throw new MaintenanceException("El tamaño recibido no coincide con el archivo seleccionado.");
        var extension=Path.GetExtension(name).ToLowerInvariant();string mime;
        if(extension==".pdf"&&bytes.AsSpan().StartsWith("%PDF-"u8))mime="application/pdf";
        else if(kind.Code!="GUIA"&&extension==".png"&&bytes.AsSpan().StartsWith(new byte[]{137,80,78,71,13,10,26,10}))mime="image/png";
        else if(kind.Code!="GUIA"&&extension is ".jpg" or ".jpeg"&&bytes.AsSpan().StartsWith(new byte[]{255,216,255})){mime="image/jpeg";extension=".jpg";}
        else throw new MaintenanceException(kind.Code=="GUIA"?"Adjunta la guía escaneada en PDF.":"Adjunta un archivo PDF, PNG o JPG válido.");
        var hash=Convert.ToHexString(SHA256.HashData(bytes));var key=(folder is null?"":folder+"/")+operation.ToString("N")+extension;
        Reference(key);
        return new(new(name,key,mime,bytes.LongLength,hash),bytes);
    }
    public static string Folder(string code,string number)
    {
        if(code=="FIRMA")return "ConfigHYT/Firma";
        var root=code switch{"OC"=>"OCCliente","GUIA"=>"Guia","CONFIRMACION"=>"ConfirmacionCliente","PRESUPUESTO"=>"Presupuestos",_=>throw new MaintenanceException("Tipo de carpeta no válido.")};
        var value=number.Trim(' ');if(value.Length is <1 or >100)throw new MaintenanceException("Revisa el número del documento para guardar su archivo.");
        return root+"/"+Regex.Replace(value,@"[^A-Za-z0-9_-]","_",RegexOptions.CultureInvariant);
    }
    public static void Reference(string key)
    {
        if(!Regex.IsMatch(key,@"\A(?:(?:(?:Presupuestos|OCCliente|Guia|ConfirmacionCliente)/[A-Za-z0-9_-]{1,100}|ConfigHYT/Firma)/)?[a-f0-9]{32}\.(pdf|png|jpg)\z",RegexOptions.CultureInvariant))throw new MaintenanceException("La referencia del archivo no es válida.");
    }
    public static void Verify(DocumentAttachment file,byte[] bytes)
    {
        Reference(file.ClaveAlmacenamiento);
        if(file.TamanoBytes is <=0 or >MaximumSize||bytes.LongLength!=file.TamanoBytes)throw new MaintenanceException("El archivo no coincide con su tamaño registrado.");
        if(!Convert.ToHexString(SHA256.HashData(bytes)).Equals(file.SHA256,StringComparison.OrdinalIgnoreCase))throw new MaintenanceException("El archivo no coincide con su huella de integridad.");
    }
}