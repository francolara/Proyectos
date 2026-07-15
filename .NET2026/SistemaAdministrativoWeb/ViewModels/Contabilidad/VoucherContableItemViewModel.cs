namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VoucherContableItemViewModel
{
    public short Item { get; set; }
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string GlosaDetalle { get; set; } = string.Empty;
    public string NumeroDocumento { get; set; } = string.Empty;
    public string TipoDocumento { get; set; } = string.Empty;
    public string Serie { get; set; } = string.Empty;
    public string Referencia { get; set; } = string.Empty;
    public string CentroCosto { get; set; } = string.Empty;
    public decimal TipoCambio { get; set; }
    public decimal Debe { get; set; }
    public decimal Haber { get; set; }
    public decimal ImporteSoles { get; set; }
    public decimal ImporteDolares { get; set; }
}
