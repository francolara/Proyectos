namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionImpuestoTabViewModel
{
    public string Titulo { get; set; } = string.Empty;
    public List<ConfiguracionImpuestoFormViewModel> Items { get; set; } = [];
}
