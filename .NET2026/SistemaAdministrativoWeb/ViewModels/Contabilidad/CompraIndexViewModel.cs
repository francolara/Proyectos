using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CompraIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public string TipoComprobanteFiltro { get; set; } = string.Empty;
    public int TotalCompras { get; set; }
    public decimal TotalImportePeriodo { get; set; }
    public decimal TotalImporteSolesPeriodo { get; set; }
    public decimal TotalImporteDolaresPeriodo { get; set; }
    public decimal PorcentajeRetencionRenta4ta { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public CompraFormViewModel Formulario { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<ProveedorDto> Proveedores { get; set; } = [];
    public List<ConfiguracionContabilizacionResumenDto> ConfiguracionesCompra { get; set; } = [];
    public List<MonedaDto> Monedas { get; set; } = [];
    public List<TipoComprobanteDto> TiposComprobante { get; set; } = [];
    public List<PlanCuentaDto> CuentasMovimiento { get; set; } = [];
    public List<TipoAfectacionIgvDto> TiposAfectacionIgv { get; set; } = [];
    public List<DetraccionSunatDto> DetraccionesSunat { get; set; } = [];
    public List<TipoPercepcionDto> TiposPercepcion { get; set; } = [];
    public List<OpcionCatalogoViewModel> TiposDocumentoIdentidad { get; set; } = [];
    public List<CompraResumenItemViewModel> Compras { get; set; } = [];
    public string ProveedorSeleccionadoTipoDocumento { get; set; } = string.Empty;
    public string ProveedorSeleccionadoNumeroDocumento { get; set; } = string.Empty;
    public string ProveedorSeleccionadoNombreLegal { get; set; } = string.Empty;
    public string ProveedorSeleccionadoTexto { get; set; } = string.Empty;
    public string ProveedorSeleccionadoContacto { get; set; } = string.Empty;
    public string ProveedorSeleccionadoCorreo { get; set; } = string.Empty;
}
