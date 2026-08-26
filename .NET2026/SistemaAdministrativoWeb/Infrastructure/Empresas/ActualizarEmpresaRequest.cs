namespace SistemaAdministrativoWeb.Infrastructure.Empresas;

public sealed class ActualizarEmpresaRequest
{
    public int IdEmpresa { get; init; }
    public string AspNetUserId { get; init; } = string.Empty;
    public string RazonSocial { get; init; } = string.Empty;
    public string? NombreComercial { get; init; }
    public string Ruc { get; init; } = string.Empty;
    public string? UsuarioRegistro { get; init; }
}
