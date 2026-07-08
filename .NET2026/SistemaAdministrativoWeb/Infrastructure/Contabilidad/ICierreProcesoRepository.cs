namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ICierreProcesoRepository
{
    Task<CierreProcesoDto?> ObtenerAsync(int idEmpresa, short anio, CancellationToken cancellationToken = default);
    Task<CierreProcesoDto> GenerarAsync(GenerarCierreProcesoRequest request, CancellationToken cancellationToken = default);
    Task EliminarAsync(int idEmpresa, short anio, string? usuarioRegistro, CancellationToken cancellationToken = default);
}
