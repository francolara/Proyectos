using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class CuentaDestinoReglaFormViewModel : IRegistroTrazable
{
    public int? IdCuentaDestinoRegla { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione la cuenta origen.")]
    public int? IdPlanCuentaOrigen { get; set; }

    public string CuentaOrigenTexto { get; set; } = string.Empty;

    [StringLength(500, ErrorMessage = "La observacion no puede exceder 500 caracteres.")]
    public string? Observacion { get; set; }

    public bool Activo { get; set; } = true;
    public string? UsuarioCreacion { get; set; }
    public DateTime? FechaRegistro { get; set; }
    public string? UsuarioActualizacion { get; set; }
    public DateTime? FechaActualizacion { get; set; }

    public List<CuentaDestinoReglaDetalleFormViewModel> Detalles { get; set; } =
    [
        new()
        {
            Orden = 1,
            Porcentaje = 100,
            Activo = true
        }
    ];
}
