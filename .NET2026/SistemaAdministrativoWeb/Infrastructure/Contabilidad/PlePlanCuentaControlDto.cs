namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PlePlanCuentaControlDto
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public string CodigoFormato { get; init; } = string.Empty;
    public string HuellaPlanContable { get; init; } = string.Empty;
    public DateTime FechaUltimaGeneracion { get; init; }
}

public sealed class PlePlanCuentaControlRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public string CodigoFormato { get; init; } = string.Empty;
    public string HuellaPlanContable { get; init; } = string.Empty;
    public string UsuarioGeneracion { get; init; } = string.Empty;
}
