using System.ComponentModel.DataAnnotations;

namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoDetalleFormViewModel
{
    public int? IdAsientoDetalle { get; set; }

    [Range(1, short.MaxValue, ErrorMessage = "El item debe ser mayor a cero.")]
    public short Item { get; set; }

    [Range(1, int.MaxValue, ErrorMessage = "Seleccione una cuenta contable.")]
    public int? IdPlanCuenta { get; set; }

    public string CuentaTexto { get; set; } = string.Empty;

    public bool RequiereCentroCostoCuenta { get; set; }

    [StringLength(300)]
    public string? GlosaDetalle { get; set; }

    [StringLength(50)]
    public string? CodigoCentroCosto { get; set; }

    public string CentroCostoTexto { get; set; } = string.Empty;

    [StringLength(150)]
    public string? TipoDocumento { get; set; }

    [StringLength(20)]
    public string? NumeroDocumento { get; set; }

    public string PersonaTexto { get; set; } = string.Empty;

    [StringLength(10)]
    public string? Serie { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal? TipoCambioLinea { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Debe { get; set; }

    [Range(typeof(decimal), "0", "999999999999")]
    public decimal Haber { get; set; }

    [StringLength(100)]
    public string? ReferenciaLinea { get; set; }
}
