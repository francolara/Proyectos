namespace GestionServicios.Application.Followup;
public sealed record DocumentAttachment(string NombreOriginal,string ClaveAlmacenamiento,string TipoContenido,long TamanoBytes,string SHA256);
public sealed record DocumentFile(long Id,DocumentAttachment Metadata);
public interface IDocumentFileRepository
{
    Task<DocumentFile?> GetAsync(DocumentKind kind,string actor,long fileId,CancellationToken ct);
}