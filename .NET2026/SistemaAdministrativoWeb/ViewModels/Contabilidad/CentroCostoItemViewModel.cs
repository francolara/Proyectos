namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CentroCostoItemViewModel
{
    public int IdCentroCosto { get; set; }
    public string CodigoCentroCosto { get; set; } = string.Empty;
    public string NombreCentroCosto { get; set; } = string.Empty;
    public bool Estado { get; set; }
}
