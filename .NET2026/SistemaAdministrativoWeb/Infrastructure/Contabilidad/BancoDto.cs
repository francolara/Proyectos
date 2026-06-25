namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BancoDto
{
    public int IdBanco { get; init; }
    public string CodigoBanco { get; init; } = string.Empty;
    public string NombreBanco { get; init; } = string.Empty;
    public bool Estado { get; init; }
}
