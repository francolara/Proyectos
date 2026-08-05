namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PleGenerationResultDto
{
    public bool Generado { get; init; }
    public string Mensaje { get; init; } = string.Empty;
    public string TokenDescarga { get; init; } = string.Empty;
    public string TokenDescargaComplementario { get; init; } = string.Empty;
    public string NombreArchivo { get; init; } = string.Empty;
    public string NombreArchivoComplementario { get; init; } = string.Empty;
    public PleConsultaResultadoDto Consulta { get; init; } = new();
}
