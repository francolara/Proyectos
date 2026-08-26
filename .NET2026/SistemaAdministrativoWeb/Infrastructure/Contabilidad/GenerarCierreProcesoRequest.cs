namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GenerarCierreProcesoRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public byte MesSaldoHasta { get; init; }
    public byte MesGeneracion { get; init; }
    public decimal TipoCambioCompra { get; init; }
    public decimal TipoCambioVenta { get; init; }
    public string? UsuarioRegistro { get; init; }
}
