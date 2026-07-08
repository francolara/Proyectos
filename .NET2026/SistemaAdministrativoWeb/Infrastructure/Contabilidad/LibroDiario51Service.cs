namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiario51Service(ILibroElectronicoRepository repository) : ILibroDiario51Service
{
    public Task<IReadOnlyCollection<LibroDiario51Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return repository.ListarLibroDiario51Async(request, cancellationToken);
    }
}
