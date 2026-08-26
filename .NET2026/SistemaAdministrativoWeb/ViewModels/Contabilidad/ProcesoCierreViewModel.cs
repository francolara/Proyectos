namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ProcesoCierreViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSaldoHastaSeleccionado { get; set; }
    public string PeriodoSaldoHasta { get; set; } = string.Empty;
    public byte MesGeneracionSeleccionado { get; set; }
    public string PeriodoGeneracion { get; set; } = string.Empty;
    public int? IdOrigenConfigurado { get; set; }
    public string OrigenConfiguradoTexto { get; set; } = string.Empty;
    public bool PuedeGenerar { get; set; }
    public bool ProcesoGenerado { get; set; }
    public bool UsaTipoCambioSbs { get; set; }
    public decimal TipoCambioCompra { get; set; }
    public decimal TipoCambioVenta { get; set; }
    public DateOnly? FechaAsiento { get; set; }
    public int? IdAsiento { get; set; }
    public int? NumeroAsiento { get; set; }
    public int TotalLineas { get; set; }
    public int TotalCuentas { get; set; }
    public int TotalAsientos { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public DateTime? FechaRegistro { get; set; }
    public string? UsuarioRegistro { get; set; }
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesContablesDisponibles { get; set; } = [];
    public List<CierreProcesoDetalleItemViewModel> Detalles { get; set; } = [];
}
