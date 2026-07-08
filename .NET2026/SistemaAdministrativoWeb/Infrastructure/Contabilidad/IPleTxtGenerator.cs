namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IPleTxtGenerator
{
    Task<byte[]> GenerarLibroDiario51Async(IReadOnlyCollection<LibroDiario51Dto> items, CancellationToken cancellationToken = default);
    Task<byte[]> GenerarLibroDiario52Async(IReadOnlyCollection<LibroDiario52Dto> items, CancellationToken cancellationToken = default);
    Task<byte[]> GenerarLibroMayor61Async(IReadOnlyCollection<LibroMayor61Dto> items, CancellationToken cancellationToken = default);
}
