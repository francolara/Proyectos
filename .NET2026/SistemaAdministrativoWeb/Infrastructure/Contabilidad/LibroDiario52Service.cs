namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class LibroDiario52Service(ILibroElectronicoRepository repository) : ILibroDiario52Service
{
    public Task<IReadOnlyCollection<LibroDiario52Dto>> ListarAsync(LibroElectronicoConsultaRequest request, CancellationToken cancellationToken = default)
    {
        return repository.ListarLibroDiario52Async(request, cancellationToken);
    }
}
