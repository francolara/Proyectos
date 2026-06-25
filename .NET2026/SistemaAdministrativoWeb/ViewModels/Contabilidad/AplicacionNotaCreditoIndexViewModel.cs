namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AplicacionNotaCreditoIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalAplicaciones { get; set; }
    public decimal TotalImporteAplicado { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<AplicacionPersonaOpcionViewModel> PersonasDisponibles { get; set; } = [];
    public List<AplicacionNotaCreditoPendienteItemViewModel> ComprobantesPendientes { get; set; } = [];
    public List<AplicacionNotaCreditoPendienteItemViewModel> NotasCreditoPendientes { get; set; } = [];
    public List<AplicacionNotaCreditoResumenItemViewModel> Aplicaciones { get; set; } = [];
    public AplicacionNotaCreditoFormViewModel Formulario { get; set; } = new();
}
