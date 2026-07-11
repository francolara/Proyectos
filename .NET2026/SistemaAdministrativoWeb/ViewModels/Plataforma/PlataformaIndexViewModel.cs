namespace SistemaAdministrativoWeb.ViewModels.Plataforma;

public sealed class PlataformaIndexViewModel
{
    public string TextoBusqueda { get; init; } = string.Empty;
    public string EstadoFiltro { get; init; } = "TODOS";
    public int TotalCuentas { get; init; }
    public int CuentasActivas { get; init; }
    public int CuentasEnPrueba { get; init; }
    public int CuentasSuspendidasOBaja { get; init; }
    public int CobrosRegistrados { get; init; }
    public int CobrosPendientesAplicacion { get; init; }
    public decimal MontoCobradoMes { get; init; }
    public IReadOnlyCollection<CuentaSuscripcionViewModel> Cuentas { get; init; } = [];
}
