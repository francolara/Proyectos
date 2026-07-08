namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPleDownloadStore
{
    string Guardar(string fileName, byte[] content);
    PleDownloadPayload? Obtener(string token);
    void Remover(string token);
}
