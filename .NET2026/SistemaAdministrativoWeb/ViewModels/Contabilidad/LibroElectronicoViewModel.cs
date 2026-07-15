using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class LibroElectronicoViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string EmpresaRuc { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string LibroElectronicoSeleccionado { get; set; } = PleLibroElectronicoCatalogo.LibroDiario51;
    public string MonedaSeleccionada { get; set; } = "PEN";
    public bool ConsultaEjecutada { get; set; }
    public bool ValidacionEjecutada { get; set; }
    public string OperacionEjecutada { get; set; } = string.Empty;
    public string? MensajeError { get; set; }
    public string? MensajeExito { get; set; }
    public string TokenDescarga { get; set; } = string.Empty;
    public bool PuedeDescargarTxt { get; set; }
    public int PaginaPreview { get; set; } = 1;
    public int PaginaHistorial { get; set; } = 1;
    public int TamanoPaginaPreview { get; set; } = 50;
    public int TamanoPaginaHistorial { get; set; } = 10;
    public PleResumenDto Resumen { get; set; } = new();
    public PleValidationResultDto Validacion { get; set; } = new();
    public PaginacionViewModel PreviewPaginacion { get; set; } = new();
    public PaginacionViewModel HistorialPaginacion { get; set; } = new();
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<OpcionCatalogoViewModel> LibrosDisponibles { get; set; } = [];
    public IReadOnlyCollection<LibroDiario51Dto> LibroDiario51Items { get; set; } = [];
    public IReadOnlyCollection<LibroDiario52Dto> LibroDiario52Items { get; set; } = [];
    public IReadOnlyCollection<LibroMayor61Dto> LibroMayor61Items { get; set; } = [];
    public IReadOnlyCollection<PleHistoryItemDto> HistorialItems { get; set; } = [];
    public bool PuedeVer { get; set; }
    public bool PuedeConsultar { get; set; }
    public bool PuedeValidar { get; set; }
    public bool PuedeGenerar { get; set; }
    public bool PuedeVerHistorial { get; set; }
    public bool PuedeDescargar { get; set; }
}
