namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoResumenItemViewModel
{
    public int IdMovimientoBanco { get; set; }
    public int? IdAsiento { get; set; }
    public int? NumeroAsiento { get; set; }
    public int NumeroMovimiento { get; set; }
    public DateOnly FechaEmision { get; set; }
    public int IdBancoConfiguracionEmpresa { get; set; }
    public string NroCuentaCorriente { get; set; } = string.Empty;
    public string CodigoMoneda { get; set; } = string.Empty;
    public string TipoMovimiento { get; set; } = string.Empty;
    public string TipoOperacion { get; set; } = string.Empty;
    public string NombrePersona { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string Glosa { get; set; } = string.Empty;
    public decimal Ingreso { get; set; }
    public decimal Egreso { get; set; }
}
