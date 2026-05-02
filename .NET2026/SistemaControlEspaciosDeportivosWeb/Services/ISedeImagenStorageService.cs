using Microsoft.AspNetCore.Http;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface ISedeImagenStorageService
{
    Task<IReadOnlyList<string>> UploadSedeImagenesAsync(int negocioId, int? sedeId, IEnumerable<IFormFile> archivos, CancellationToken cancellationToken = default);
    Task<string?> UploadLogoNegocioAsync(int negocioId, IFormFile? archivo, CancellationToken cancellationToken = default);
    Task<string?> UploadBannerPublicoAsync(IFormFile? archivo, CancellationToken cancellationToken = default);
    Task<string?> UploadBannerPublicoFlexibleAsync(IFormFile? archivo, CancellationToken cancellationToken = default);
    Task<string?> UploadBannerPublicoMobileAsync(IFormFile? archivo, CancellationToken cancellationToken = default);
    Task<string?> UploadBannerAnuncioAsync(IFormFile? archivo, bool esHorizontal, CancellationToken cancellationToken = default);
    Task<(bool Ok, string Mensaje)> ProbarConexionAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<string>> DeleteSedeImagenesAsync(IEnumerable<string> urls, CancellationToken cancellationToken = default);
    Task<(byte[] Contenido, string ContentType)?> ObtenerImagenVisualizacionAsync(string? url, CancellationToken cancellationToken = default);
}
