namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class RegistroComprasViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public string CodigoPersona { get; set; } = string.Empty;
    public string NumeroComprobante { get; set; } = string.Empty;
    public bool ConsultaEjecutada { get; set; }
    public string? MensajeError { get; set; }
    public decimal TotalBaseImponibleGravada { get; set; }
    public decimal TotalIgvGravado { get; set; }
    public decimal TotalBaseImponibleGasto { get; set; }
    public decimal TotalIgvGasto { get; set; }
    public decimal TotalBaseImponibleSinCredito { get; set; }
    public decimal TotalIgvSinCredito { get; set; }
    public decimal TotalExonerado { get; set; }
    public decimal TotalInafecto { get; set; }
    public decimal TotalOtrosTributos { get; set; }
    public decimal TotalIcbper { get; set; }
    public decimal TotalRetencion { get; set; }
    public decimal TotalDetraccion { get; set; }
    public decimal TotalPercepcion { get; set; }
    public decimal TotalImporte { get; set; }
    public List<short> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
    public List<RegistroComprasItemViewModel> Resultados { get; set; } = [];
}
