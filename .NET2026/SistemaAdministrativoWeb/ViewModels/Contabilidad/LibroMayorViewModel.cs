namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class LibroMayorViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string PeriodoConsulta { get; set; } = string.Empty;
    public string CuentaDesde { get; set; } = string.Empty;
    public string CuentaHasta { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public bool ConsultaEjecutada { get; set; }
    public string? MensajeError { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public decimal TotalDebeDolares { get; set; }
    public decimal TotalHaberDolares { get; set; }
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<LibroMayorItemViewModel> Resultados { get; set; } = [];
}
