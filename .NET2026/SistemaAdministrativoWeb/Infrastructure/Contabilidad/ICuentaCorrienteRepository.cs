using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICuentaCorrienteRepository
{
    Task<PagedResult<BancoConfiguracionEmpresaDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, string? textoBusqueda, int numeroPagina, int tamanoPagina, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<BancoConfiguracionEmpresaDto>> ListarPorEmpresaAsync(int idEmpresa, bool soloActivos = false, CancellationToken cancellationToken = default);
    Task<BancoConfiguracionEmpresaDto?> ObtenerPorIdAsync(int idEmpresa, int idBancoConfiguracionEmpresa, CancellationToken cancellationToken = default);
    Task<BancoConfiguracionEmpresaDto> GuardarAsync(GuardarBancoConfiguracionEmpresaRequest request, CancellationToken cancellationToken = default);
}
