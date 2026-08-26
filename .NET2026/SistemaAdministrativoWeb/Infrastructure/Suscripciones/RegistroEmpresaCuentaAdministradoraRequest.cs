namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class RegistroEmpresaCuentaAdministradoraRequest
{
    public int IdCuentaAdministradora { get; init; }
    public string AspNetUserId { get; init; } = string.Empty;
    public string CodigoEmpresa { get; init; } = string.Empty;
    public string RazonSocial { get; init; } = string.Empty;
    public string? NombreComercial { get; init; }
    public string Ruc { get; init; } = string.Empty;
    public bool EsEmpresaPredeterminada { get; init; }
    public string? UsuarioRegistro { get; init; }
}
