namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class UsuarioPerfilRequest
{
    public string AspNetUserId { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string? Telefono { get; init; }
    public string? CorreoReferencia { get; init; }
    public string? UsuarioRegistro { get; init; }
}
