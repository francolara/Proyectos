using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalAsientos { get; set; }
    public decimal TotalDebePeriodo { get; set; }
    public decimal TotalHaberPeriodo { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public AsientoFormViewModel Formulario { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<OrigenDto> OrigenesManual { get; set; } = [];
    public List<MonedaDto> Monedas { get; set; } = [];
    public List<PlanCuentaDto> CuentasMovimiento { get; set; } = [];
    public List<AsientoResumenItemViewModel> Asientos { get; set; } = [];
}
