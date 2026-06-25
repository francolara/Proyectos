using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class TransferenciaCuentaIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalTransferencias { get; set; }
    public decimal TotalImporteEmisor { get; set; }
    public decimal TotalImporteReceptor { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<BancoConfiguracionEmpresaDto> CuentasCorrientesDisponibles { get; set; } = [];
    public List<CajaBancoOperacionViewModel> OperacionesEmisor { get; set; } = [];
    public List<CajaBancoOperacionViewModel> OperacionesReceptor { get; set; } = [];
    public List<TransferenciaCuentaResumenItemViewModel> Transferencias { get; set; } = [];
    public TransferenciaCuentaFormViewModel Formulario { get; set; } = new();
}
