namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarConfiguracionContabilizacionDetalleRequest
{
    public short Orden { get; init; }
    public string ComponenteContable { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string NaturalezaMovimiento { get; init; } = string.Empty;
    public bool Activo { get; init; }
}
