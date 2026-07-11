namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class PlataformaCobrosViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public string EstadoPagoFiltro { get; init; } = "TODOS";
    public int TotalCobros { get; init; }
    public int CobrosPagados { get; init; }
    public int CobrosPendientes { get; init; }
    public decimal TotalMontoPagado { get; init; }
    public decimal TotalMontoPendiente { get; init; }
    public IReadOnlyCollection<PlataformaCobroItemViewModel> Cobros { get; init; } = [];
}
