namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ITipoAfectacionIgvRepository
{
    Task<IReadOnlyCollection<TipoAfectacionIgvDto>> ListarActivosAsync(CancellationToken cancellationToken = default);
}
