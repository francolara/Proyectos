namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionContabilizacionResumenItemViewModel
{
    public int IdConfiguracionContabilizacion { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public string EscenarioOperacion { get; init; } = string.Empty;
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public bool GeneraAsientoAutomatico { get; init; }
    public bool UsaTipoCambio { get; init; }
    public bool Activo { get; init; }
    public int CantidadComponentes { get; init; }
}
