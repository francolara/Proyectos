using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IBancoRepository
{
    Task<PagedResult<BancoDto>> ListarPaginadoAsync(string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = true, CancellationToken cancellationToken = default);
}
