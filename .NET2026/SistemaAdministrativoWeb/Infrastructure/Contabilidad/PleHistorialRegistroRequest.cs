namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleHistorialRegistroRequest
{
    public int IdEmpresa { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public string CodigoLibro { get; init; } = string.Empty;
    public string CodigoFormato { get; init; } = string.Empty;
    public string NombreArchivo { get; init; } = string.Empty;
    public int CantidadRegistros { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string Observaciones { get; init; } = string.Empty;
    public string UsuarioGeneracion { get; init; } = string.Empty;
    public string CodigoFormatoComplementario { get; init; } = string.Empty;
    public string NombreArchivoComplementario { get; init; } = string.Empty;
    public int CantidadRegistrosComplementario { get; init; }
    public string HuellaPlanContable { get; init; } = string.Empty;
    public string PlanContableSnapshot { get; init; } = string.Empty;
}
