namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class VentaDetalleDto
{
    public int IdVentaDetalle { get; init; }
    public int IdVenta { get; init; }
    public short Item { get; init; }
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public int IdTipoAfectacionIGV { get; init; }
    public string CodigoAfectacionIGV { get; init; } = string.Empty;
    public string NombreAfectacionIGV { get; init; } = string.Empty;
    public string Descripcion { get; init; } = string.Empty;
    public decimal Cantidad { get; init; }
    public decimal ValorUnitario { get; init; }
    public decimal ImporteBruto { get; init; }
}
