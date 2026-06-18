using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICuentaDestinoReglaRepository
{
    Task<IReadOnlyCollection<CuentaDestinoReglaResumenDto>> ListarPorEmpresaAsync(int idEmpresa, short ejercicio, CancellationToken cancellationToken = default);
    Task<PagedResult<CuentaDestinoReglaResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<CuentaDestinoReglaDto?> ObtenerAsync(int idCuentaDestinoRegla, CancellationToken cancellationToken = default);
    Task<int> GuardarAsync(GuardarCuentaDestinoReglaRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idCuentaDestinoRegla, CancellationToken cancellationToken = default);
}
