namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class OrigenItemViewModel
{
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string ModuloOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public bool Estado { get; init; }
}
