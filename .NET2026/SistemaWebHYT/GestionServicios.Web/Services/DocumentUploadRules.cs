using System.Security.Cryptography;
using System.Text.RegularExpressions;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.Services;
public sealed record ValidatedDocumentUpload(DocumentAttachment Metadata,byte[] Bytes);
public static class DocumentUploadRules
{
    public const long MaximumSize=15*1024*1024;
    public static async Task<ValidatedDocumentUpload> ValidateAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct)
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
        var hash=Convert.ToHexString(SHA256.HashData(bytes));var key=operation.ToString("N")+extension;
        return new(new(name,key,mime,bytes.LongLength,hash),bytes);
    }
    public static void Reference(string key)
    {
        if(!Regex.IsMatch(key,@"^[a-f0-9]{32}\.(pdf|png|jpg)$",RegexOptions.CultureInvariant))throw new MaintenanceException("La referencia del archivo no es válida.");
    }
    public static void Verify(DocumentAttachment file,byte[] bytes)
    {
        Reference(file.ClaveAlmacenamiento);
        if(file.TamanoBytes is <=0 or >MaximumSize||bytes.LongLength!=file.TamanoBytes)throw new MaintenanceException("El archivo no coincide con su tamaño registrado.");
        if(!Convert.ToHexString(SHA256.HashData(bytes)).Equals(file.SHA256,StringComparison.OrdinalIgnoreCase))throw new MaintenanceException("El archivo no coincide con su huella de integridad.");
    }
}