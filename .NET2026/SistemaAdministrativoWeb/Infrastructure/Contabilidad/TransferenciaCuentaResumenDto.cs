namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TransferenciaCuentaResumenDto
{
    public Guid IdTransferenciaCuenta { get; init; }
    public int IdMovimientoBancoEmisor { get; init; }
    public int? IdAsientoEmisor { get; init; }
    public int NumeroMovimientoEmisor { get; init; }
    public int? NumeroAsientoEmisor { get; init; }
    public int IdBancoConfiguracionEmpresaEmisor { get; init; }
    public string CuentaCorrienteEmisor { get; init; } = string.Empty;
    public string MonedaEmisor { get; init; } = string.Empty;
    public string OperacionEmisor { get; init; } = string.Empty;
    public DateOnly FechaEmisionEmisor { get; init; }
    public decimal TipoCambioEmisor { get; init; }
    public string NumeroOperacionEmisor { get; init; } = string.Empty;
    public decimal ImporteEmisor { get; init; }
    public string GlosaEmisor { get; init; } = string.Empty;
    public int IdMovimientoBancoReceptor { get; init; }
    public int? IdAsientoReceptor { get; init; }
    public int NumeroMovimientoReceptor { get; init; }
    public int? NumeroAsientoReceptor { get; init; }
    public int IdBancoConfiguracionEmpresaReceptor { get; init; }
    public string CuentaCorrienteReceptor { get; init; } = string.Empty;
    public string MonedaReceptor { get; init; } = string.Empty;
    public string OperacionReceptor { get; init; } = string.Empty;
    public DateOnly FechaEmisionReceptor { get; init; }
    public decimal TipoCambioReceptor { get; init; }
    public string NumeroOperacionReceptor { get; init; } = string.Empty;
    public decimal ImporteReceptor { get; init; }
    public string GlosaReceptor { get; init; } = string.Empty;
}
