using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICompraRepository
{
    Task<IReadOnlyCollection<CompraResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CompraDetraccionPendienteDto>> ListarDetraccionesPendientesPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default);
    Task<PagedResult<CompraResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<CompraDto?> ObtenerAsync(int idCompra, CancellationToken cancellationToken = default);
    Task<GuardarCompraResultDto> GuardarAsync(GuardarCompraRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idCompra, int idEmpresa, CancellationToken cancellationToken = default);
}
