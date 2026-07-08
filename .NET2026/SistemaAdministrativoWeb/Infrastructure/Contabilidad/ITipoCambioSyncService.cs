namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ITipoCambioSyncService
{
    Task<TipoCambioDto?> SincronizarFechaAsync(int idCuentaAdministradora, DateOnly fecha, string idMoneda, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task<TipoCambioPeriodoSyncResult> SincronizarPeriodoAsync(int idCuentaAdministradora, short anio, byte mes, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
