using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CajaBancoIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public int? IdBancoConfiguracionEmpresaSeleccionada { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public decimal SaldoInicial { get; set; }
    public decimal IngresosMes { get; set; }
    public decimal EgresosMes { get; set; }
    public decimal SaldoFinal { get; set; }
    public int TotalMovimientos { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<BancoConfiguracionEmpresaDto> CuentasCorrientesDisponibles { get; set; } = [];
    public List<CajaBancoOperacionViewModel> OperacionesIngreso { get; set; } = [];
    public List<CajaBancoOperacionViewModel> OperacionesEgreso { get; set; } = [];
    public List<OpcionCatalogoViewModel> TiposDocumentoDisponibles { get; set; } = [];
    public List<CajaBancoResumenItemViewModel> Movimientos { get; set; } = [];
    public CajaBancoFormViewModel Formulario { get; set; } = new();
}
