namespace SistemaAdministrativoWeb.ViewModels.Panel;

public sealed class PanelSeriePeriodoViewModel
{
    public string Periodo { get; init; } = string.Empty;
    public int Registros { get; init; }
    public decimal ImporteTotalPen { get; init; }
    public decimal ImporteTotalUsd { get; init; }
}
