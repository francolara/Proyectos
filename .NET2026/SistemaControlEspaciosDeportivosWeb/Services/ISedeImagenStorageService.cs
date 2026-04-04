using Microsoft.AspNetCore.Http;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface ISedeImagenStorageService
{
    Task<IReadOnlyList<string>> UploadSedeImagenesAsync(int negocioId, int? sedeId, IEnumerable<IFormFile> archivos, CancellationToken cancellationToken = default);
    Task<(bool Ok, string Mensaje)> ProbarConexionAsync(CancellationToken cancellationToken = default);
    Task<IReadOnlyList<string>> DeleteSedeImagenesAsync(IEnumerable<string> urls, CancellationToken cancellationToken = default);
    Task<(byte[] Contenido, string ContentType)?> ObtenerImagenVisualizacionAsync(string? url, CancellationToken cancellationToken = default);
}
