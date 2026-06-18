namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class MonedaDto
{
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public string NombreMoneda { get; init; } = string.Empty;
    public string SimboloMoneda { get; init; } = string.Empty;
    public bool EsMonedaBase { get; init; }
    public bool Estado { get; init; }
}
