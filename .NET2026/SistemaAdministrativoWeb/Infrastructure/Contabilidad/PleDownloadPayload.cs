namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleDownloadPayload
{
    public string Token { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public byte[] Content { get; init; } = [];
    public string ContentType { get; init; } = "application/octet-stream";
    public DateTimeOffset FechaExpiracion { get; init; }
}
