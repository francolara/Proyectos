namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TransferenciaCuentaResumenItemViewModel
{
    public Guid IdTransferenciaCuenta { get; set; }
    public int IdMovimientoBancoEmisor { get; set; }
    public int NumeroMovimientoEmisor { get; set; }
    public int? NumeroAsientoEmisor { get; set; }
    public string CuentaCorrienteEmisor { get; set; } = string.Empty;
    public string MonedaEmisor { get; set; } = string.Empty;
    public string OperacionEmisor { get; set; } = string.Empty;
    public DateOnly FechaEmisionEmisor { get; set; }
    public string NumeroOperacionEmisor { get; set; } = string.Empty;
    public decimal ImporteEmisor { get; set; }
    public string GlosaEmisor { get; set; } = string.Empty;
    public int NumeroMovimientoReceptor { get; set; }
    public int? NumeroAsientoReceptor { get; set; }
    public string CuentaCorrienteReceptor { get; set; } = string.Empty;
    public string MonedaReceptor { get; set; } = string.Empty;
    public string OperacionReceptor { get; set; } = string.Empty;
    public DateOnly FechaEmisionReceptor { get; set; }
    public string NumeroOperacionReceptor { get; set; } = string.Empty;
    public decimal ImporteReceptor { get; set; }
    public string GlosaReceptor { get; set; } = string.Empty;
}
