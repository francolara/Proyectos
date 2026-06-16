namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class RegistroCuentaAdministradoraConEmpresaRequest
{
    public string AspNetUserId { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string? Telefono { get; init; }
    public string CorreoReferencia { get; init; } = string.Empty;
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CodigoEmpresa { get; init; } = string.Empty;
    public string RazonSocial { get; init; } = string.Empty;
    public string? NombreComercial { get; init; }
    public string Ruc { get; init; } = string.Empty;
    public int DiasPrueba { get; init; } = 30;
    public string? UsuarioRegistro { get; init; }
}
