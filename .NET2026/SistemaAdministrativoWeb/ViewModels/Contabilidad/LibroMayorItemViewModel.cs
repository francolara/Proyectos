namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class LibroMayorItemViewModel
{
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string CodigoOrigen { get; set; } = string.Empty;
    public string NombreOrigen { get; set; } = string.Empty;
    public string Periodo { get; set; } = string.Empty;
    public int NumeroAsiento { get; set; }
    public short Item { get; set; }
    public DateOnly? FechaEmision { get; set; }
    public string TipoDocumento { get; set; } = string.Empty;
    public string Serie { get; set; } = string.Empty;
    public string Referencia { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string NombreAuxiliar { get; set; } = string.Empty;
    public string Glosa { get; set; } = string.Empty;
    public decimal TipoCambio { get; set; }
    public decimal Debe { get; set; }
    public decimal Haber { get; set; }
    public decimal DebeDolares { get; set; }
    public decimal HaberDolares { get; set; }
    public decimal SaldoInicial { get; set; }
    public decimal SaldoInicialDolares { get; set; }
    public decimal SaldoAcumulado { get; set; }
    public decimal SaldoAcumuladoDolares { get; set; }
    public bool EsSaldoInicial { get; set; }
}
