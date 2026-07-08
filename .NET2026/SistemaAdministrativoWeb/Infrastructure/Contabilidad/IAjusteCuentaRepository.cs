namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAjusteCuentaRepository
{
    Task<AjusteCuentaProcesoDto?> ObtenerAsync(int idEmpresa, string periodo, CancellationToken cancellationToken = default);
    Task<AjusteCuentaProcesoDto> GenerarAsync(GenerarAjusteCuentaProcesoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, string periodo, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
