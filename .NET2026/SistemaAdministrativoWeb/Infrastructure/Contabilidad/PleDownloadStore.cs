using Microsoft.Extensions.Caching.Memory;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleDownloadStore(IMemoryCache memoryCache) : IPleDownloadStore
{
    private static readonly TimeSpan TiempoVida = TimeSpan.FromMinutes(15);

    public string Guardar(string fileName, byte[] content, string contentType = "text/plain; charset=utf-8")
    {
        var token = Guid.NewGuid().ToString("N");
        var payload = new PleDownloadPayload
        {
            Token = token,
            FileName = fileName,
            Content = content,
            ContentType = contentType,
            FechaExpiracion = DateTimeOffset.Now.Add(TiempoVida)
        };

        memoryCache.Set(token, payload, payload.FechaExpiracion);
        return token;
    }

    public PleDownloadPayload? Obtener(string token)
    {
        return string.IsNullOrWhiteSpace(token)
            ? null
            : memoryCache.Get<PleDownloadPayload>(token.Trim());
    }

    public void Remover(string token)
    {
        if (!string.IsNullOrWhiteSpace(token))
        {
            memoryCache.Remove(token.Trim());
        }
    }
}
