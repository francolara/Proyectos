namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public interface ICuentaAdministradoraRepository
{
    Task GuardarPerfilUsuarioAsync(UsuarioPerfilRequest request, CancellationToken cancellationToken = default);
    Task<RegistroCuentaAdministradoraConEmpresaResult> RegistrarCuentaConEmpresaAsync(RegistroCuentaAdministradoraConEmpresaRequest request, CancellationToken cancellationToken = default);
    Task<RegistroEmpresaCuentaAdministradoraResult> RegistrarEmpresaCuentaAsync(RegistroEmpresaCuentaAdministradoraRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaSuscripcionResumenDto>> ListarCuentasSuscripcionAsync(CancellationToken cancellationToken = default);
    Task ActualizarSuscripcionCuentaAsync(ActualizarSuscripcionCuentaRequest request, CancellationToken cancellationToken = default);
    Task<ContextoSuscripcionEmpresaDto?> ObtenerContextoSuscripcionPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default);
}
