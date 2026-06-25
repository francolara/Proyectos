namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class OperacionBancariaDto
{
    public string IdOpeBancaria { get; init; } = string.Empty;
    public string IdTipoOpeBancaria { get; init; } = string.Empty;
    public string TipoMovimiento { get; init; } = string.Empty;
    public string TipoOperacion { get; init; } = string.Empty;
}
