namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface ILibroMayorRepository
{
    Task<IReadOnlyCollection<LibroMayorItemDto>> ListarAsync(LibroMayorRequest request, CancellationToken cancellationToken = default);
}
