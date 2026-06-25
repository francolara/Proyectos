namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class OrigenIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalOrigenes { get; set; }
    public int TotalActivos { get; set; }
    public int TotalManual { get; set; }
    public bool PuedeCargarDefault { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public OrigenFormViewModel Formulario { get; set; } = new();
    public List<OrigenItemViewModel> Origenes { get; set; } = [];
}
