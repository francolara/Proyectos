namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class PlataformaIndexViewModel
{
    public int TotalCuentas { get; init; }
    public int CuentasActivas { get; init; }
    public int CuentasEnPrueba { get; init; }
    public int CuentasSuspendidasOBaja { get; init; }
    public IReadOnlyCollection<CuentaSuscripcionViewModel> Cuentas { get; init; } = [];
}
