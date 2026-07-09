namespace SistemaAdministrativoWeb.ViewModels.Panel;

public sealed class PanelDistribucionItemViewModel
{
    public string Titulo { get; init; } = string.Empty;
    public int Valor { get; init; }
    public decimal Porcentaje { get; init; }
    public string Color { get; init; } = "#2f8cff";
    public string Descripcion { get; init; } = string.Empty;
}
