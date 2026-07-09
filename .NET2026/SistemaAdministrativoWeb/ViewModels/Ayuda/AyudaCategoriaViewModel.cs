namespace SistemaAdministrativoWeb.ViewModels.Ayuda;

public sealed class AyudaCategoriaViewModel
{
    public string Clave { get; init; } = string.Empty;
    public string Titulo { get; init; } = string.Empty;
    public string Icono { get; init; } = "bi-collection";
    public string Descripcion { get; init; } = string.Empty;
    public IReadOnlyCollection<AyudaModuloViewModel> Modulos { get; init; } = [];
}
