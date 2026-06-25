namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CentroCostoIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalCentrosCosto { get; set; }
    public int TotalActivos { get; set; }
    public int TotalInactivos { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public CentroCostoFormViewModel Formulario { get; set; } = new();
    public List<CentroCostoItemViewModel> CentrosCosto { get; set; } = [];
}
