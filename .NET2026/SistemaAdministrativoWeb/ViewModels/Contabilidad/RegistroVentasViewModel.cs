namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class RegistroVentasViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string CodigoPersona { get; set; } = string.Empty;
    public bool ConsultaEjecutada { get; set; }
    public string? MensajeError { get; set; }
    public decimal TotalBaseImponible { get; set; }
    public decimal TotalDescuento { get; set; }
    public decimal TotalExonerado { get; set; }
    public decimal TotalInafecto { get; set; }
    public decimal TotalIgv { get; set; }
    public decimal TotalIsc { get; set; }
    public decimal TotalOtrosTributos { get; set; }
    public decimal TotalIcbper { get; set; }
    public decimal TotalRedondeo { get; set; }
    public decimal TotalImporte { get; set; }
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<RegistroVentasItemViewModel> Resultados { get; set; } = [];
}
