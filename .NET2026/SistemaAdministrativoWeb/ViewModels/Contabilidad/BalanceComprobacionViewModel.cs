namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class BalanceComprobacionViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte PeriodoDesdeSeleccionado { get; set; }
    public byte PeriodoHastaSeleccionado { get; set; }
    public string MonedaSeleccionada { get; set; } = "PEN";
    public byte GradoSeleccionado { get; set; } = 1;
    public bool TodasLasCuentas { get; set; } = true;
    public string CuentaDesde { get; set; } = string.Empty;
    public string CuentaHasta { get; set; } = string.Empty;
    public bool FiltrarGrado { get; set; } = true;
    public bool ConsultaEjecutada { get; set; }
    public string? MensajeError { get; set; }
    public decimal TotalDebAnt { get; set; }
    public decimal TotalHabAnt { get; set; }
    public decimal TotalDebMes { get; set; }
    public decimal TotalHabMes { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public decimal TotalResultadoDebe { get; set; }
    public decimal TotalResultadoHaber { get; set; }
    public decimal TotalActivo { get; set; }
    public decimal TotalPasivo { get; set; }
    public decimal TotalPerdidaNaturaleza { get; set; }
    public decimal TotalGananciaNaturaleza { get; set; }
    public decimal TotalPerdidaFuncion { get; set; }
    public decimal TotalGananciaFuncion { get; set; }
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> PeriodosDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> MonedasDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> CuentasDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> GradosDisponibles { get; set; } = [];
    public List<BalanceComprobacionItemViewModel> Resultados { get; set; } = [];
}
