namespace SistemaAdministrativoWeb.ViewModels.Ayuda;

public sealed class AyudaIndexViewModel
{
    public string CategoriaSeleccionadaClave { get; init; } = string.Empty;
    public string ModuloSeleccionadoClave { get; init; } = string.Empty;
    public string ModuloSeleccionadoTitulo { get; init; } = string.Empty;
    public string? ModuloSolicitado { get; init; }
    public int TotalPreguntas { get; init; }
    public IReadOnlyCollection<AyudaCategoriaViewModel> Categorias { get; init; } = [];
}
