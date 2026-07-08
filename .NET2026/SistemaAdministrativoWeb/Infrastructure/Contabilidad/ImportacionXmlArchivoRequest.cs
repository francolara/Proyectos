namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ImportacionXmlArchivoRequest
{
    public string NombreArchivo { get; init; } = string.Empty;
    public byte[] Contenido { get; init; } = [];
}
