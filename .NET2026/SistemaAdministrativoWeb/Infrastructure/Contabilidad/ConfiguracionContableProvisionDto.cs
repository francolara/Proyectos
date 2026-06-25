namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContableProvisionDto
{
    public int? IdConfiguracionContabilizacion { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public int? IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public bool GeneraAsientoAutomatico { get; init; } = true;
    public bool Activo { get; init; } = true;
}
