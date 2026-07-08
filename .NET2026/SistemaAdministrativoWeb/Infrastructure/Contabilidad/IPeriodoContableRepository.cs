namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPeriodoContableRepository
{
    Task<PeriodoContableEstadoDto?> ObtenerAsync(int idEmpresa, string periodo, CancellationToken cancellationToken = default);
    Task<PeriodoContableEstadoDto> GuardarAsync(GuardarPeriodoContableEstadoRequest request, CancellationToken cancellationToken = default);
}
