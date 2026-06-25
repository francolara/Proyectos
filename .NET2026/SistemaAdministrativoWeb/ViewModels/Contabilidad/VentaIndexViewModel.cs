using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VentaIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalVentas { get; set; }
    public decimal TotalImportePeriodo { get; set; }
    public decimal TotalImporteSolesPeriodo { get; set; }
    public decimal TotalImporteDolaresPeriodo { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public VentaFormViewModel Formulario { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<ClienteDto> Clientes { get; set; } = [];
    public List<ConfiguracionContabilizacionResumenDto> ConfiguracionesVenta { get; set; } = [];
    public List<MonedaDto> Monedas { get; set; } = [];
    public List<TipoComprobanteDto> TiposComprobante { get; set; } = [];
    public List<PlanCuentaDto> CuentasMovimiento { get; set; } = [];
    public List<TipoAfectacionIgvDto> TiposAfectacionIgv { get; set; } = [];
    public List<OpcionCatalogoViewModel> TiposDocumentoIdentidad { get; set; } = [];
    public List<VentaResumenItemViewModel> Ventas { get; set; } = [];
    public string ClienteSeleccionadoTipoDocumento { get; set; } = string.Empty;
    public string ClienteSeleccionadoNumeroDocumento { get; set; } = string.Empty;
    public string ClienteSeleccionadoNombreLegal { get; set; } = string.Empty;
    public string ClienteSeleccionadoTexto { get; set; } = string.Empty;
    public string ClienteSeleccionadoTelefono { get; set; } = string.Empty;
    public string ClienteSeleccionadoCorreo { get; set; } = string.Empty;
}
