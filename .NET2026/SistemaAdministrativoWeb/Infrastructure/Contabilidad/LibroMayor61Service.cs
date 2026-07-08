namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroMayor61Service(ILibroElectronicoRepository repository) : ILibroMayor61Service
{
    public Task<IReadOnlyCollection<LibroMayor61Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return repository.ListarLibroMayor61Async(request, cancellationToken);
    }
}
