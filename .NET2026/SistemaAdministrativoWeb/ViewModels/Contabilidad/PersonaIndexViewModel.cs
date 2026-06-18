namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PersonaIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public string TipoPersonaFiltro { get; set; } = string.Empty;
    public bool SoloClientes { get; set; }
    public bool SoloProveedores { get; set; }
    public int TotalPersonas { get; set; }
    public int TotalClientes { get; set; }
    public int TotalProveedores { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public PersonaFormViewModel Formulario { get; set; } = new();
    public List<PersonaItemViewModel> Personas { get; set; } = [];
    public List<OpcionCatalogoViewModel> TiposPersona { get; set; } = [];
    public List<OpcionCatalogoViewModel> TiposDocumento { get; set; } = [];
    public List<OpcionCatalogoViewModel> Departamentos { get; set; } = [];
    public List<OpcionCatalogoViewModel> Provincias { get; set; } = [];
    public List<OpcionCatalogoViewModel> Distritos { get; set; } = [];
}
