using System.Net;
using System.Text.RegularExpressions;
using Amazon.Runtime;
using Amazon.S3;
using Amazon.S3.Model;
using GestionServicios.Application.Followup;
using GestionServicios.Application.Maintenance;
using Microsoft.Extensions.Options;
namespace GestionServicios.Web.Services;
public sealed class CloudflareR2Settings
{
    public string AccountId {get;set;}="";
    public string BucketName {get;set;}="hyt-documentos";
    public string AccessKeyId {get;set;}="";
    public string SecretAccessKey {get;set;}="";
}
public interface IR2ObjectClient
{
    Task<bool> PutIfAbsentAsync(DocumentAttachment metadata,byte[] bytes,CancellationToken ct);
    Task<byte[]?> GetAsync(string key,CancellationToken ct);
}
public sealed class CloudflareR2ObjectClient(IOptions<CloudflareR2Settings> options):IR2ObjectClient
{
    private AmazonS3Client Client()
    {
        var settings=options.Value;
        if(!Regex.IsMatch(settings.AccountId,@"^[a-fA-F0-9]{32}$")||!Regex.IsMatch(settings.BucketName,@"^[a-z0-9][a-z0-9-]{1,61}[a-z0-9]$")||string.IsNullOrWhiteSpace(settings.AccessKeyId)||string.IsNullOrWhiteSpace(settings.SecretAccessKey))
            throw new MaintenanceException("El almacenamiento de documentos todavía no está configurado.");
        return new AmazonS3Client(new BasicAWSCredentials(settings.AccessKeyId,settings.SecretAccessKey),new AmazonS3Config{ServiceURL=$"https://{settings.AccountId}.r2.cloudflarestorage.com",AuthenticationRegion="auto",ForcePathStyle=true,Timeout=TimeSpan.FromSeconds(30),MaxErrorRetry=2});
    }
    public async Task<bool> PutIfAbsentAsync(DocumentAttachment metadata,byte[] bytes,CancellationToken ct)
    {
        DocumentUploadRules.Verify(metadata,bytes);using var client=Client();using var stream=new MemoryStream(bytes,false);
        var request=new PutObjectRequest{BucketName=options.Value.BucketName,Key=metadata.ClaveAlmacenamiento,InputStream=stream,ContentType=metadata.TipoContenido,IfNoneMatch="*",DisablePayloadSigning=true,DisableDefaultChecksumValidation=true};
        request.Metadata["hyt-sha256"]=metadata.SHA256;
        try{await client.PutObjectAsync(request,ct);return true;}
        catch(AmazonS3Exception error)when(error.StatusCode==HttpStatusCode.PreconditionFailed||error.StatusCode==HttpStatusCode.Conflict){return false;}
        catch(Exception error)when(error is AmazonS3Exception or HttpRequestException||error is OperationCanceledException&&!ct.IsCancellationRequested){throw new MaintenanceException("No se pudo guardar el archivo en Cloudflare. Revisa el registro antes de reintentar.");}
    }
    public async Task<byte[]?> GetAsync(string key,CancellationToken ct)
    {
        DocumentUploadRules.Reference(key);using var client=Client();
        try
        {
            using var response=await client.GetObjectAsync(new GetObjectRequest{BucketName=options.Value.BucketName,Key=key},ct);
            if(response.ContentLength is <=0 or >DocumentUploadRules.MaximumSize)throw new MaintenanceException("El archivo almacenado no tiene un tamaño válido.");
            using var buffer=new MemoryStream();var block=new byte[81920];int count;
            while((count=await response.ResponseStream.ReadAsync(block,ct))>0){if(buffer.Length+count>DocumentUploadRules.MaximumSize)throw new MaintenanceException("El archivo supera el tamaño permitido.");await buffer.WriteAsync(block.AsMemory(0,count),ct);}
            return buffer.ToArray();
        }
        catch(AmazonS3Exception error)when(error.ErrorCode=="NoSuchKey"){return null;}
        catch(Exception error)when(error is AmazonS3Exception or HttpRequestException or IOException||error is OperationCanceledException&&!ct.IsCancellationRequested){throw new MaintenanceException("No se pudo recuperar el archivo desde Cloudflare. Vuelve a intentarlo.");}
    }
}
public sealed class R2DocumentFileStore(IR2ObjectClient objects):IDocumentFileStore
{
    public async Task<DocumentAttachment> SaveAsync(IFormFile upload,DocumentKind kind,Guid operation,CancellationToken ct)
    {
        var validated=await DocumentUploadRules.ValidateAsync(upload,kind,operation,ct);
        if(!await objects.PutIfAbsentAsync(validated.Metadata,validated.Bytes,ct))
        {
            var prior=await objects.GetAsync(validated.Metadata.ClaveAlmacenamiento,ct);
            if(prior is null)throw new MaintenanceException("No se pudo confirmar el archivo anterior. Reintenta la misma operación.");
            DocumentUploadRules.Verify(validated.Metadata,prior);
        }
        return validated.Metadata;
    }
    public async Task<byte[]?> ReadAsync(DocumentAttachment file,CancellationToken ct)
    {
        DocumentUploadRules.Reference(file.ClaveAlmacenamiento);var bytes=await objects.GetAsync(file.ClaveAlmacenamiento,ct);
        if(bytes is not null)DocumentUploadRules.Verify(file,bytes);return bytes;
    }
}