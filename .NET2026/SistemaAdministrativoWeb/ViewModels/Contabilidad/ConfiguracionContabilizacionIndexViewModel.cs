namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContabilizacionIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalConfiguraciones { get; set; }
    public int TotalActivas { get; set; }
    public int TotalAutomaticas { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public ConfiguracionContabilizacionFormViewModel Formulario { get; set; } = new();
    public List<OrigenDto> Origenes { get; set; } = [];
    public List<PlanCuentaDto> CuentasMovimiento { get; set; } = [];
    public List<ConfiguracionProvisionOperacionViewModel> Provisiones { get; set; } = [];
    public List<ConfiguracionDocumentoFormViewModel> Documentos { get; set; } = [];
    public List<ConfiguracionImpuestoFormViewModel> Impuestos { get; set; } = [];
    public List<ConfiguracionParametroContableFormViewModel> Parametros { get; set; } = [];
    public List<ConfiguracionContabilizacionResumenItemViewModel> Configuraciones { get; set; } = [];
}
