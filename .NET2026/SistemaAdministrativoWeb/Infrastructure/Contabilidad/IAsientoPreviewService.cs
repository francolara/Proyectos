namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAsientoPreviewService
{
    Task<AsientoPreviewResultDto> PrevisualizarAsync(int idEmpresa, AsientoPreviewRequest request, CancellationToken cancellationToken = default);
}
