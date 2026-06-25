namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TransferenciaCuentaFormViewModel
{
    public TransferenciaCuentaSeccionFormViewModel Emisor { get; set; } = new();
    public TransferenciaCuentaSeccionFormViewModel Receptor { get; set; } = new();
}
