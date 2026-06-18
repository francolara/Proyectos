using SistemaAdministrativoWeb.Infrastructure.Contabilidad;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaDestinoReglaIndexViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short EjercicioActual { get; set; }
    public string TextoBusqueda { get; set; } = string.Empty;
    public int TotalReglas { get; set; }
    public int TotalActivas { get; set; }
    public int TotalTramos { get; set; }
    public decimal PorcentajeConfigurado { get; set; }
    public PaginacionViewModel Paginacion { get; set; } = new();
    public CuentaDestinoReglaFormViewModel Formulario { get; set; } = new();
    public List<PlanCuentaDto> CuentasMovimiento { get; set; } = [];
    public List<CuentaDestinoReglaResumenItemViewModel> Reglas { get; set; } = [];
}
