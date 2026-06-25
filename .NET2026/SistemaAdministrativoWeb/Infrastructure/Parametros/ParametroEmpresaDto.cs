namespace SistemaAdministrativoWeb.Infrastructure.Parametros;

public sealed class ParametroEmpresaDto
{
    public int IdParametroEmpresa { get; init; }
    public int IdEmpresa { get; init; }
    public string TipoParametro { get; init; } = string.Empty;
    public string CodigoParametro { get; init; } = string.Empty;
    public string ValorParametro { get; init; } = string.Empty;
    public string DescripcionParametro { get; init; } = string.Empty;
    public DateTime? FecIni { get; init; }
    public DateTime? FecFin { get; init; }
    public bool Activo { get; init; }
}
