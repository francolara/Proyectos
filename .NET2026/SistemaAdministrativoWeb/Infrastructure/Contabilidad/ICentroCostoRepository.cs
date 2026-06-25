using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICentroCostoRepository
{
    Task<IReadOnlyCollection<CentroCostoDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<PagedResult<CentroCostoDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<CentroCostoDto> GuardarAsync(GuardarCentroCostoRequest request, CancellationToken cancellationToken = default);
}
