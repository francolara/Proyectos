namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContabilizacionDetalleDto
{
    public int IdConfiguracionContabilizacionDetalle { get; init; }
    public int IdConfiguracionContabilizacion { get; init; }
    public short Orden { get; init; }
    public string ComponenteContable { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string NaturalezaMovimiento { get; init; } = string.Empty;
    public bool Activo { get; init; }
}
