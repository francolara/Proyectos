namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PlanCuentaIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public byte? NivelCuentaFiltro { get; set; }
    public int TotalCuentas { get; set; }
    public int TotalMovimiento { get; set; }
    public int TotalActivas { get; set; }
    public bool PuedeCargarDefault { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public PlanCuentaFormViewModel Formulario { get; set; } = new();
    public List<PlanCuentaItemViewModel> Cuentas { get; set; } = [];
    public List<PlanCuentaItemViewModel> CuentasPadre { get; set; } = [];
    public List<OpcionCatalogoViewModel> Monedas { get; set; } = [];
}
