using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAsientoRepository
{
    Task<IReadOnlyCollection<AsientoResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, bool soloManual = false, CancellationToken cancellationToken = default);
    Task<PagedResult<AsientoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloManual = false, CancellationToken cancellationToken = default);
    Task<AsientoDto?> ObtenerAsync(int idAsiento, CancellationToken cancellationToken = default);
    Task<GuardarAsientoManualResultDto> GuardarManualAsync(GuardarAsientoManualRequest request, CancellationToken cancellationToken = default);
}
