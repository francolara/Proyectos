using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IVentaRepository
{
    Task<IReadOnlyCollection<VentaResumenDto>> ListarPorEmpresaAsync(int idEmpresa, string? periodo = null, CancellationToken cancellationToken = default);
    Task<PagedResult<VentaResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short ejercicio, byte mes, string? textoBusqueda, string? tipoComprobante, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<VentaDto?> ObtenerAsync(int idVenta, CancellationToken cancellationToken = default);
    Task<GuardarVentaResultDto> GuardarAsync(GuardarVentaRequest request, CancellationToken cancellationToken = default);
    Task<ImportarVentaXmlResultDto> ImportarXmlAsync(ImportarVentaXmlRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idVenta, int idEmpresa, CancellationToken cancellationToken = default);
}
