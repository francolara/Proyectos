using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaCorrienteIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalCuentasCorrientes { get; set; }
    public int TotalActivas { get; set; }
    public int TotalBancos { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public List<MonedaDto> Monedas { get; set; } = [];
    public CuentaCorrienteFormViewModel Formulario { get; set; } = new();
    public List<CuentaCorrienteItemViewModel> CuentasCorrientes { get; set; } = [];
}
