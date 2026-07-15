namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoEstadoCuentaItemViewModel
{
    public int IdMovimientoBanco { get; set; }
    public int NumeroMovimiento { get; set; }
    public DateOnly FechaEmision { get; set; }
    public string TipoMovimiento { get; set; } = string.Empty;
    public string TipoOperacion { get; set; } = string.Empty;
    public string NombrePersona { get; set; } = string.Empty;
    public string NumeroOperacion { get; set; } = string.Empty;
    public string Glosa { get; set; } = string.Empty;
    public decimal Ingreso { get; set; }
    public decimal Egreso { get; set; }
    public decimal SaldoAcumulado { get; set; }
}
