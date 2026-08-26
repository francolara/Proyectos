namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class LibroDiarioViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte PeriodoSeleccionado { get; set; }
    public string PeriodoConsulta { get; set; } = string.Empty;
    public string MonedaSeleccionada { get; set; } = "PEN";
    public string ModoSeleccionado { get; set; } = "A";
    public string CuentaDesde { get; set; } = string.Empty;
    public string CuentaHasta { get; set; } = string.Empty;
    public string OrigenDesde { get; set; } = string.Empty;
    public string OrigenHasta { get; set; } = string.Empty;
    public bool ConsultaEjecutada { get; set; }
    public string? MensajeError { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public decimal TotalDebeDolares { get; set; }
    public decimal TotalHaberDolares { get; set; }
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> PeriodosDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> MonedasDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> ModosDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> OrigenesDisponibles { get; set; } = [];
    public List<LibroDiarioItemViewModel> Resultados { get; set; } = [];
}
