using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPlanCuentaRepository
{
    Task<IReadOnlyCollection<PlanCuentaDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloMovimiento = false, CancellationToken cancellationToken = default);
    Task<PagedResult<PlanCuentaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloMovimiento = false, CancellationToken cancellationToken = default);
    Task<PlanCuentaDto> GuardarAsync(GuardarPlanCuentaRequest request, CancellationToken cancellationToken = default);
}
