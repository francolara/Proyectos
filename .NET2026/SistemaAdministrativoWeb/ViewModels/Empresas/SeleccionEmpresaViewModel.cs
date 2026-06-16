using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Empresas;

public sealed class SeleccionEmpresaViewModel
{
    [Required]
    public int IdEmpresaSeleccionada { get; set; }

    public List<EmpresaDisponibleViewModel> Empresas { get; set; } = [];
}
