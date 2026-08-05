using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroElectronicoRepository
{
    Task<IReadOnlyCollection<LibroDiario51Dto>> ListarLibroDiario51Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<LibroDiario52Dto>> ListarLibroDiario52Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<LibroMayor61Dto>> ListarLibroMayor61Async(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default);
    Task<PagedResult<PleHistoryItemDto>> ListarHistorialAsync(int idEmpresa, short anio, byte mes, string? libroElectronico, int numeroPagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task RegistrarHistorialAsync(PleHistorialRegistroRequest request, CancellationToken cancellationToken = default);
    Task<PlePlanCuentaControlDto?> ObtenerControlPlanAsync(int idEmpresa, short anio, string codigoFormato, CancellationToken cancellationToken = default);
    Task GuardarControlPlanAsync(PlePlanCuentaControlRequest request, CancellationToken cancellationToken = default);
    Task<PlePresentacionContextoDto> ObtenerContextoPresentacionAsync(int idEmpresa, short anio, byte mes, string codigoLibro, CancellationToken cancellationToken = default);
    Task ActualizarPresentacionAsync(PlePresentacionUpdateRequest request, CancellationToken cancellationToken = default);
}
