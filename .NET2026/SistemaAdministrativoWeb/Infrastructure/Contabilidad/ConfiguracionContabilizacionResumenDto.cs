namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class ConfiguracionContabilizacionResumenDto
{
    public int IdConfiguracionContabilizacion { get; init; }
    public int IdEmpresa { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public bool GeneraAsientoAutomatico { get; init; }
    public bool UsaTipoCambio { get; init; }
    public bool Activo { get; init; }
    public int CantidadComponentes { get; init; }
}
