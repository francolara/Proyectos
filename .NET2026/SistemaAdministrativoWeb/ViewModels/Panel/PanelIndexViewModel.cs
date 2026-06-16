namespace SistemaAdministrativoWeb.ViewModels.Panel;

public sealed class PanelIndexViewModel
{
    public int IdEmpresa { get; init; }
    public string NombreEmpresa { get; init; } = string.Empty;
    public string? Ruc { get; init; }
    public string EstadoPanel { get; init; } = string.Empty;
    public string ResumenVigencia { get; init; } = string.Empty;
    public IReadOnlyCollection<PanelModuloViewModel> Modulos { get; init; } = [];
}
