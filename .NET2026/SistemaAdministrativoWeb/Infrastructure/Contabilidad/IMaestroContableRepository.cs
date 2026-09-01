using SistemaAdministrativoWeb.Infrastructure.Data;

namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IMaestroContableRepository
{
    Task<PagedResult<PlanCuentaMaestroDto>> ListarPlanCuentaAsync(string? textoBusqueda, byte? nivelCuenta, bool soloMovimiento, bool soloActivos, int pagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<PlanCuentaMaestroDto?> ObtenerPlanCuentaAsync(int idPlanCuentaMaestro, CancellationToken cancellationToken = default);
    Task<int> GuardarPlanCuentaAsync(GuardarPlanCuentaMaestroRequest request, CancellationToken cancellationToken = default);
    Task EliminarPlanCuentaAsync(int idPlanCuentaMaestro, CancellationToken cancellationToken = default);
    Task<PagedResult<CuentaDestinoMaestroResumenDto>> ListarCuentasDestinoAsync(string? textoBusqueda, int pagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<CuentaDestinoMaestroDto?> ObtenerCuentaDestinoAsync(int idCuentaDestinoReglaMaestro, CancellationToken cancellationToken = default);
    Task<int> GuardarCuentaDestinoAsync(GuardarCuentaDestinoMaestroRequest request, CancellationToken cancellationToken = default);
    Task EliminarCuentaDestinoAsync(int idCuentaDestinoReglaMaestro, CancellationToken cancellationToken = default);
    Task<AsignacionesMaestroDto> ListarAsignacionesAsync(CancellationToken cancellationToken = default);
    Task GuardarAsignacionAsync(GuardarAsignacionMaestroRequest request, CancellationToken cancellationToken = default);
    Task<PagedResult<OrigenMaestroDto>> ListarOrigenesAsync(string? textoBusqueda, bool soloActivos, int pagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task<OrigenMaestroDto?> ObtenerOrigenAsync(int idOrigenMaestro, CancellationToken cancellationToken = default);
    Task<int> GuardarOrigenAsync(GuardarOrigenMaestroRequest request, CancellationToken cancellationToken = default);
    Task EliminarOrigenAsync(int idOrigenMaestro, CancellationToken cancellationToken = default);
    Task<PagedResult<ConfiguracionContabilizacionMaestroDto>> ListarConfiguracionesAsync(string? textoBusqueda, int pagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task GuardarOrigenConfiguracionAsync(int idConfiguracionContabilizacionMaestro, string codigoOrigen, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<ValidacionMaestroIssueDto>> ValidarAsync(CancellationToken cancellationToken = default);
}
