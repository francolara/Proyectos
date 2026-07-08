namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ProcesoCerrarPeriodoViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string PeriodoConsulta { get; set; } = string.Empty;
    public short AnioSeleccionado { get; set; }
    public byte MesSeleccionado { get; set; }
    public bool Cerrado { get; set; }
    public string EstadoTexto => Cerrado ? "Cerrado" : "Abierto";
    public DateTime? FechaCierre { get; set; }
    public string? UsuarioCierre { get; set; }
    public DateTime? FechaApertura { get; set; }
    public string? UsuarioApertura { get; set; }
    public List<int> AniosDisponibles { get; set; } = [];
    public List<MesOpcionViewModel> MesesDisponibles { get; set; } = [];
}
