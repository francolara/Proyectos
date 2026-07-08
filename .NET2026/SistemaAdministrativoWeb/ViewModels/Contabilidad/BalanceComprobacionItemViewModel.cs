namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class BalanceComprobacionItemViewModel
{
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string ColBalance { get; set; } = string.Empty;
    public byte GradoCuenta { get; set; }
    public decimal DebAnt { get; set; }
    public decimal HabAnt { get; set; }
    public decimal DebMes { get; set; }
    public decimal HabMes { get; set; }
    public decimal Debe { get; set; }
    public decimal Haber { get; set; }
    public decimal ResultadoDebe { get; set; }
    public decimal ResultadoHaber { get; set; }
    public decimal Activo { get; set; }
    public decimal Pasivo { get; set; }
    public decimal PerdidaNaturaleza { get; set; }
    public decimal GananciaNaturaleza { get; set; }
    public decimal PerdidaFuncion { get; set; }
    public decimal GananciaFuncion { get; set; }
}
