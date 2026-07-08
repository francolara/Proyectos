namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public interface IMigoPadronApiClient
{
    Task<MigoRucResultDto?> ConsultarRucAsync(string ruc, CancellationToken cancellationToken = default);
    Task<MigoDniResultDto?> ConsultarDniAsync(string dni, CancellationToken cancellationToken = default);
    Task<MigoCpeResultDto?> ValidarCpeAsync(MigoCpeRequestDto request, CancellationToken cancellationToken = default);
}
