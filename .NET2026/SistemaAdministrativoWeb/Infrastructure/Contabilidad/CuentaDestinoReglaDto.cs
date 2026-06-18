namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CuentaDestinoReglaDto
{
    public int IdCuentaDestinoRegla { get; init; }
    public int IdEmpresa { get; init; }
    public short Ejercicio { get; init; }
    public int IdPlanCuentaOrigen { get; init; }
    public string CodigoCuentaOrigen { get; init; } = string.Empty;
    public string NombreCuentaOrigen { get; init; } = string.Empty;
    public bool Activo { get; init; }
    public string? Observacion { get; init; }
    public List<CuentaDestinoReglaDetalleDto> Detalles { get; init; } = [];
}
