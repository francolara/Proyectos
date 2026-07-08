using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TipoCambioIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public int IdCuentaAdministradora { get; set; }
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public int TotalTiposCambio { get; set; }
    public int TotalActivos { get; set; }
    public int TotalInactivos { get; set; }
    public TipoCambioFormViewModel Formulario { get; set; } = new();
    public List<TipoCambioItemViewModel> TiposCambio { get; set; } = [];
    public List<MonedaDto> Monedas { get; set; } = [];
    public List<OpcionCatalogoViewModel> Fuentes { get; set; } = [];
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
}
