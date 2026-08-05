namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPleDownloadStore
{
    string Guardar(string fileName, byte[] content, string contentType = "text/plain; charset=utf-8");
    PleDownloadPayload? Obtener(string token);
    void Remover(string token);
}
