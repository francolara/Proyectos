namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class VoucherContableViewModel
{
    public int IdEmpresa { get; set; }
    public string EmpresaNombre { get; set; } = string.Empty;
    public string RucEmpresa { get; set; } = string.Empty;
    public int IdAsiento { get; set; }
    public string Periodo { get; set; } = string.Empty;
    public int NumeroAsiento { get; set; }
    public string CodigoOrigen { get; set; } = string.Empty;
    public string NombreOrigen { get; set; } = string.Empty;
    public string Glosa { get; set; } = string.Empty;
    public string Moneda { get; set; } = string.Empty;
    public decimal TipoCambio { get; set; }
    public DateOnly FechaEmision { get; set; }
    public string ReferenciaExterna { get; set; } = string.Empty;
    public string Observacion { get; set; } = string.Empty;
    public bool MuestraColumnaDolares { get; set; }
    public List<VoucherContableItemViewModel> Detalles { get; set; } = [];
    public decimal TotalDebe => Detalles.Sum(x => x.Debe);
    public decimal TotalHaber => Detalles.Sum(x => x.Haber);
    public decimal TotalSoles => Detalles.Sum(x => x.ImporteSoles);
    public decimal TotalDolares => Detalles.Sum(x => x.ImporteDolares);
}
