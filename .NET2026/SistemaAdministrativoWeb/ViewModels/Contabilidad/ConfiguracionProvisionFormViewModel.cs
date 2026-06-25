using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionProvisionFormViewModel
{
    [Required]
    public string ModuloOperacion { get; set; } = string.Empty;

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione un origen contable.")]
    public int? IdOrigen { get; set; }

    public string OrigenTexto { get; set; } = string.Empty;
    public bool GeneraAsientoAutomatico { get; set; } = true;
    public bool Activo { get; set; } = true;
}
