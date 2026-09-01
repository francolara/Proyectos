namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public interface ICuentaAdministradoraRepository
{
    Task GuardarPerfilUsuarioAsync(UsuarioPerfilRequest request, CancellationToken cancellationToken = default);
    Task<RegistroCuentaAdministradoraConEmpresaResult> RegistrarCuentaConEmpresaAsync(RegistroCuentaAdministradoraConEmpresaRequest request, CancellationToken cancellationToken = default);
    Task<RegistroEmpresaCuentaAdministradoraResult> RegistrarEmpresaCuentaAsync(RegistroEmpresaCuentaAdministradoraRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaSuscripcionResumenDto>> ListarCuentasSuscripcionAsync(CancellationToken cancellationToken = default);
    Task<CuentaSuscripcionPaginaDto> ListarCuentasSuscripcionPaginadasAsync(string? textoBusqueda, string estadoFiltro, int pagina, int tamanoPagina, CancellationToken cancellationToken = default);
    Task SincronizarVencimientoSuscripcionCuentaAsync(int idCuentaAdministradora, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task ActualizarSuscripcionCuentaAsync(ActualizarSuscripcionCuentaRequest request, CancellationToken cancellationToken = default);
    Task ActivarContratoCuentaAsync(ActivarContratoCuentaRequest request, CancellationToken cancellationToken = default);
    Task RegistrarPagoSuscripcionCuentaAsync(RegistrarPagoSuscripcionCuentaRequest request, CancellationToken cancellationToken = default);
    Task ConfirmarPagoSuscripcionCuentaAsync(int idCuentaAdministradora, int idCuentaAdministradoraSuscripcionPago, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaSuscripcionMovimientoDto>> ListarMovimientosSuscripcionCuentaAsync(int idCuentaAdministradora, int top = 20, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaSuscripcionPagoDto>> ListarPagosSuscripcionCuentaAsync(int idCuentaAdministradora, int top = 20, CancellationToken cancellationToken = default);
    Task<ContextoSuscripcionEmpresaDto?> ObtenerContextoSuscripcionPorEmpresaAsync(int idEmpresa, CancellationToken cancellationToken = default);
    Task<ContextoLoginUsuarioDto?> ObtenerContextoLoginUsuarioAsync(string aspNetUserId, CancellationToken cancellationToken = default);
    Task<CuentaAdministradoraConfiguracionDto?> ObtenerConfiguracionCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default);
    Task GuardarConfiguracionCuentaAdministradoraAsync(GuardarConfiguracionCuentaAdministradoraRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaAdministradoraEmpresaDto>> ListarEmpresasCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<CuentaAdministradoraUsuarioDto>> ListarUsuariosCuentaAdministradoraAsync(int idCuentaAdministradora, CancellationToken cancellationToken = default);
    Task AsignarUsuarioCuentaAdministradoraAsync(AsignarUsuarioCuentaAdministradoraRequest request, CancellationToken cancellationToken = default);
    Task AsignarUsuarioEmpresaAsync(AsignarUsuarioEmpresaRequest request, CancellationToken cancellationToken = default);
    Task DesactivarUsuarioCuentaAdministradoraAsync(string aspNetUserId, int idCuentaAdministradora, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task DesactivarUsuarioEmpresaAsync(string aspNetUserId, int idEmpresa, string? usuarioRegistro, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UsuarioCuentaEmpresaAsignadaDto>> ListarEmpresasUsuarioCuentaAdministradoraAsync(int idCuentaAdministradora, string aspNetUserId, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UsuarioCuentaPermisoDto>> ListarPermisosUsuarioCuentaAsync(int idUsuarioCuentaAdministradora, CancellationToken cancellationToken = default);
    Task GuardarUsuarioCuentaPermisoAsync(GuardarUsuarioCuentaPermisoRequest request, CancellationToken cancellationToken = default);
    Task GuardarPermisosUsuarioCuentaAsync(GuardarPermisosUsuarioCuentaRequest request, CancellationToken cancellationToken = default);
    Task<IReadOnlyCollection<UsuarioCuentaPermisoDto>> ListarPermisosUsuarioEmpresaAsync(int idUsuarioEmpresa, CancellationToken cancellationToken = default);
    Task GuardarUsuarioEmpresaPermisoAsync(GuardarUsuarioEmpresaPermisoRequest request, CancellationToken cancellationToken = default);
    Task GuardarPermisosUsuarioEmpresaAsync(GuardarPermisosUsuarioEmpresaRequest request, CancellationToken cancellationToken = default);
}
