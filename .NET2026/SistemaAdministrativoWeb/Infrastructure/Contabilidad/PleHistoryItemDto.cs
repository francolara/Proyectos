namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleHistoryItemDto
{
    public int IdLibroElectronicoGeneracion { get; init; }
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
    public DateTime FechaGeneracion { get; init; }
    public string UsuarioGeneracion { get; init; } = string.Empty;
    public string CodigoFormatoComplementario { get; init; } = string.Empty;
    public string NombreArchivoComplementario { get; init; } = string.Empty;
    public int CantidadRegistrosComplementario { get; init; }
    public bool PlanPresentado { get; init; }
    public DateTime? FechaPresentacion { get; init; }
    public string UsuarioPresentacion { get; init; } = string.Empty;
    public int TotalRegistros { get; init; }
}
