namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IXmlProvisionImportService
{
    Task<ImportacionXmlResultadoDto> ImportarComprasAsync(int idEmpresa, IReadOnlyCollection<ImportacionXmlArchivoRequest> archivos, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task<ImportacionXmlResultadoDto> ImportarVentasAsync(int idEmpresa, IReadOnlyCollection<ImportacionXmlArchivoRequest> archivos, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
