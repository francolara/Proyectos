using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAplicacionNotaCreditoRepository
{
    Task<PagedResult<AplicacionNotaCreditoResumenDto>> ListarPaginadoPorEmpresaAsync(int idEmpresa, short anio, byte mes, string? textoBusqueda, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<AplicacionNotaCreditoPendienteDto>> ListarPendientesPorPersonaAsync(int idEmpresa, string moduloOperacion, int idPersona, CancellationToken cancellationToken = default);
    Task<GuardarAplicacionNotaCreditoResultDto> GuardarAsync(GuardarAplicacionNotaCreditoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idAplicacionNotaCredito, int idEmpresa, CancellationToken cancellationToken = default);
}
