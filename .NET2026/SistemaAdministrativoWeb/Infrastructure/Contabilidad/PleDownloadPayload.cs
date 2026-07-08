namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleDownloadPayload
{
    public string Token { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public byte[] Content { get; init; } = [];
    public DateTimeOffset FechaExpiracion { get; init; }
}
