namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionParametroContableFormViewModel
{
    public int IdParametroEmpresa { get; set; }
    public string TipoParametro { get; set; } = string.Empty;
    public string CodigoParametro { get; set; } = string.Empty;
    public string DescripcionParametro { get; set; } = string.Empty;
    public string ValorParametro { get; set; } = string.Empty;
    public string CuentaTexto { get; set; } = string.Empty;
    public bool Activo { get; set; } = true;
}
