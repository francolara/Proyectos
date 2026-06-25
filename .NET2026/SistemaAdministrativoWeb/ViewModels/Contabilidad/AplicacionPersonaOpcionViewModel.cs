namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionPersonaOpcionViewModel
{
    public int IdPersona { get; init; }
    public string CodigoAuxiliar { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string NombreCompleto { get; init; } = string.Empty;
    public string Texto => $"{CodigoAuxiliar} | {NombreCompleto} ({NumeroDocumento})";
}
