namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionProvisionOperacionViewModel
{
    public string ModuloOperacion { get; set; } = string.Empty;
    public string Titulo { get; set; } = string.Empty;
    public string Descripcion { get; set; } = string.Empty;
    public string Resumen { get; set; } = string.Empty;
    public string Icono { get; set; } = "bi-diagram-3";
    public string SufijoHtml { get; set; } = string.Empty;
    public ConfiguracionProvisionFormViewModel Formulario { get; set; } = new();
}
