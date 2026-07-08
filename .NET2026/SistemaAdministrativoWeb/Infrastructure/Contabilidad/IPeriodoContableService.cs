namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPeriodoContableService
{
    Task<PeriodoContableEstadoDto> ObtenerEstadoAsync(int idEmpresa, short anio, byte mes, CancellationToken cancellationToken = default);
    Task<bool> EstaCerradoAsync(int idEmpresa, short anio, byte mes, CancellationToken cancellationToken = default);
    Task<PeriodoContableEstadoDto> GuardarEstadoAsync(int idEmpresa, short anio, byte mes, bool cerrado, string? usuarioRegistro, CancellationToken cancellationToken = default);
    string ConstruirMensajeBloqueo(short anio, byte mes);
}
