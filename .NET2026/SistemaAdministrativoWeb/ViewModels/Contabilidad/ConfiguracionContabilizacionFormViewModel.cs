using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class ConfiguracionContabilizacionFormViewModel
{
    public int? IdConfiguracionContabilizacion { get; set; }

    [Required(ErrorMessage = "Seleccione el modulo.")]
    [StringLength(10)]
    public string ModuloOperacion { get; set; } = "COM";

    [Required(ErrorMessage = "Seleccione el escenario.")]
    [StringLength(20)]
    public string EscenarioOperacion { get; set; } = "MERCADERIA";

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione el origen contable.")]
    public int? IdOrigen { get; set; }

    [Required(ErrorMessage = "Ingrese una descripcion.")]
    [StringLength(200)]
    public string Descripcion { get; set; } = string.Empty;

    public bool GeneraAsientoAutomatico { get; set; } = true;
    public bool UsaTipoCambio { get; set; } = true;
    public bool Activo { get; set; } = true;

    public List<ConfiguracionContabilizacionDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Orden = 1,
            ComponenteContable = "BRUTO",
            NaturalezaMovimiento = "D",
            Activo = true
        },
        new()
        {
            Orden = 2,
            ComponenteContable = "IGV",
            NaturalezaMovimiento = "D",
            Activo = true
        },
        new()
        {
            Orden = 3,
            ComponenteContable = "TOTAL",
            NaturalezaMovimiento = "H",
            Activo = true
        }
    ];
}
