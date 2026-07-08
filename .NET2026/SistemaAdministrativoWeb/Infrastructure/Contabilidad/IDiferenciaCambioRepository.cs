namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IDiferenciaCambioRepository
{
    Task<DiferenciaCambioProcesoDto?> ObtenerAsync(int idEmpresa, string periodo, CancellationToken cancellationToken = default);
    Task<DiferenciaCambioProcesoDto> GenerarAsync(GenerarDiferenciaCambioProcesoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, string periodo, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
