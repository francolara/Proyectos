namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TipoAfectacionIgvDto
{
    public int IdTipoAfectacionIGV { get; init; }
    public string CodigoSunat { get; init; } = string.Empty;
    public string NombreAfectacion { get; init; } = string.Empty;
    public bool Estado { get; init; }
}
