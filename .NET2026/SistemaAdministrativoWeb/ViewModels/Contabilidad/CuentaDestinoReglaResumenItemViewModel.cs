namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaDestinoReglaResumenItemViewModel
{
    public int IdCuentaDestinoRegla { get; init; }
    public short Ejercicio { get; init; }
    public int IdPlanCuentaOrigen { get; init; }
    public string CodigoCuentaOrigen { get; init; } = string.Empty;
    public string NombreCuentaOrigen { get; init; } = string.Empty;
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public int CantidadTramos { get; init; }
    public decimal PorcentajeTotal { get; init; }
}
