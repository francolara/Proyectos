namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoEstadoCuentaViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string PeriodoConsulta { get; set; } = string.Empty;
    public int IdBancoConfiguracionEmpresa { get; set; }
    public string NombreBanco { get; set; } = string.Empty;
    public string NroCuentaCorriente { get; set; } = string.Empty;
    public string Titular { get; set; } = string.Empty;
    public string CodigoMoneda { get; set; } = string.Empty;
    public string NombreMoneda { get; set; } = string.Empty;
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string PeriodoSaldoInicial { get; set; } = string.Empty;
    public decimal SaldoInicialDebe { get; set; }
    public decimal SaldoInicialHaber { get; set; }
    public decimal SaldoInicial { get; set; }
    public decimal IngresosMes { get; set; }
    public decimal EgresosMes { get; set; }
    public decimal SaldoFinal { get; set; }
    public List<CajaBancoEstadoCuentaItemViewModel> Movimientos { get; set; } = [];
}
