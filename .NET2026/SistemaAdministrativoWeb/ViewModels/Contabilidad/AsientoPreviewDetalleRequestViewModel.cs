namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class AsientoPreviewDetalleRequestViewModel
{
    public short Item { get; set; }
    public string? Descripcion { get; set; }
    public decimal Cantidad { get; set; }
    public decimal ValorUnitario { get; set; }
    public decimal ImporteBruto { get; set; }
}
