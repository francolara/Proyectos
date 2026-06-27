namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionParametroTabViewModel
{
    public string Titulo { get; set; } = string.Empty;
    public List<ConfiguracionParametroContableFormViewModel> Items { get; set; } = [];
}
