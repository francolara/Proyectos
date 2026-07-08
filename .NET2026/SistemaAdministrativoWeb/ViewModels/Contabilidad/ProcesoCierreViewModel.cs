namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ProcesoCierreViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public int? IdOrigenConfigurado { get; set; }
    public string OrigenConfiguradoTexto { get; set; } = string.Empty;
    public bool PuedeGenerar { get; set; }
    public bool ProcesoGenerado { get; set; }
    public bool UsaTipoCambioSbs { get; set; }
    public decimal TipoCambioCompra { get; set; }
    public decimal TipoCambioVenta { get; set; }
    public bool ProcesarGananciasPerdidas { get; set; } = true;
    public bool ProcesarInventarios { get; set; } = true;
    public bool GananciasPerdidasGenerado { get; set; }
    public bool InventariosGenerado { get; set; }
    public DateOnly? FechaAsiento { get; set; }
    public int TotalCuentas { get; set; }
    public int TotalAsientos { get; set; }
    public decimal TotalDebe { get; set; }
    public decimal TotalHaber { get; set; }
    public DateTime? FechaRegistro { get; set; }
    public string? UsuarioRegistro { get; set; }
    public List<int> AniosDisponibles { get; set; } = [];
    public List<CierreProcesoDetalleItemViewModel> DetallesGananciasPerdidas { get; set; } = [];
    public List<CierreProcesoDetalleItemViewModel> DetallesInventarios { get; set; } = [];
}
