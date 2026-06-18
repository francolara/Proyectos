using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IConfiguracionContabilizacionRepository
{
    Task<IReadOnlyCollection<ConfiguracionContabilizacionResumenDto>> ListarPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default);
    Task<PagedResult<ConfiguracionContabilizacionResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<ConfiguracionContabilizacionDto?> ObtenerAsync(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default);
    Task<int> GuardarAsync(GuardarConfiguracionContabilizacionRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default);
}
