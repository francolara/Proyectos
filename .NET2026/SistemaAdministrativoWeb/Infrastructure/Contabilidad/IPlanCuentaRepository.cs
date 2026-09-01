using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPlanCuentaRepository
{
    Task<IReadOnlyCollection<PlanCuentaDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloMovimiento = false, CancellationToken cancellationToken = default);
    Task<PagedResult<PlanCuentaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, byte? nivelCuenta, int numeroPagina, int tamanoPagina, bool soloMovimiento = false, bool soloUltimoNivel = false, CancellationToken cancellationToken = default);
    Task<PlanCuentaDto> GuardarAsync(GuardarPlanCuentaRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, int idPlanCuenta, CancellationToken cancellationToken = default);
    Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default, int? idEmpresaBase = null);
}
