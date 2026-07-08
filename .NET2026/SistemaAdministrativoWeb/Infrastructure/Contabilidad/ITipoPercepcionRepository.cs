namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ITipoPercepcionRepository
{
    Task<IReadOnlyCollection<TipoPercepcionDto>> ListarActivasAsync(CancellationToken cancellationToken = default);
}
