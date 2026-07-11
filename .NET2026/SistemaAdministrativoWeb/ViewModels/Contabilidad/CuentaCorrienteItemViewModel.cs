namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaCorrienteItemViewModel
{
    public int IdBancoConfiguracionEmpresa { get; set; }
    public string CodigoBanco { get; set; } = string.Empty;
    public string NombreBanco { get; set; } = string.Empty;
    public string NroCuentaCorriente { get; set; } = string.Empty;
    public string Titular { get; set; } = string.Empty;
    public int? IdMoneda { get; set; }
    public string MonedaTexto { get; set; } = string.Empty;
    public int IdPlanCuenta { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string PeriodoSaldoInicial { get; set; } = string.Empty;
    public decimal SaldoInicialDebe { get; set; }
    public decimal SaldoInicialHaber { get; set; }
    public bool Activo { get; set; }
    public DateTime FechaRegistro { get; set; }
    public string UsuarioRegistro { get; set; } = string.Empty;
}
