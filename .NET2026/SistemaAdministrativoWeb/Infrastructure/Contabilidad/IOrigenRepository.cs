using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IOrigenRepository
{
    Task<IReadOnlyCollection<OrigenDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<PagedResult<OrigenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<OrigenDto> GuardarAsync(GuardarOrigenRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, int idOrigen, CancellationToken cancellationToken = default);
    Task CargarDefaultAsync(int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
