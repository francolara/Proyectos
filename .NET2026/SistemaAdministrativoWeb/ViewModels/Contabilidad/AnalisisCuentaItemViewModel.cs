namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AnalisisCuentaItemViewModel
{
    public string CodigoCuenta { get; set; } = string.Empty;
    public string NombreCuenta { get; set; } = string.Empty;
    public string Auxiliar { get; set; } = string.Empty;
    public string NombreAuxiliar { get; set; } = string.Empty;
    public string TipoDocumento { get; set; } = string.Empty;
    public string Serie { get; set; } = string.Empty;
    public string NumeroReferencia { get; set; } = string.Empty;
    public string Periodo { get; set; } = string.Empty;
    public string Comprobante { get; set; } = string.Empty;
    public string GlosaDetalle { get; set; } = string.Empty;
    public DateOnly? FechaEmision { get; set; }
    public decimal TipoCambio { get; set; }
    public decimal Debe { get; set; }
    public decimal Haber { get; set; }
    public decimal DebeDolares { get; set; }
    public decimal HaberDolares { get; set; }
}
