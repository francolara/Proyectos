namespace SistemaAdministrativoWeb.ViewModels.Panel;

public sealed class PanelModuloViewModel
{
    public string Titulo { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public string Icono { get; init; } = "bi-grid";
    public string Url { get; init; } = "#";
    public bool Habilitado { get; init; }
}
