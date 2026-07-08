namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IAperturaProcesoRepository
{
    Task<AperturaProcesoDto?> ObtenerAsync(int idEmpresa, short anioApertura, CancellationToken cancellationToken = default);
    Task<AperturaProcesoDto> GenerarAsync(GenerarAperturaProcesoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, short anioApertura, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
