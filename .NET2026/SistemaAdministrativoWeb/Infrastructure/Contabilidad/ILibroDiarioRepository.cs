namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroDiarioRepository
{
    Task<IReadOnlyCollection<LibroDiarioItemDto>> ListarAsync(LibroDiarioRequest request, CancellationToken cancellationToken = default);
}
