namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ITipoCambioRepository
{
    Task<IReadOnlyCollection<TipoCambioDto>> ListarPorCuentaAdministradoraAsync(int idCuentaAdministradora, short anio, byte mes, CancellationToken cancellationToken = default);
    Task<TipoCambioDto?> ObtenerAsync(int idTipoCambio, int idCuentaAdministradora, CancellationToken cancellationToken = default);
    Task<TipoCambioDto?> ObtenerPorFechaMonedaAsync(int idCuentaAdministradora, DateOnly fecha, string idMoneda, CancellationToken cancellationToken = default);
    Task<TipoCambioDto> GuardarAsync(GuardarTipoCambioRequest request, CancellationToken cancellationToken = default);
}
