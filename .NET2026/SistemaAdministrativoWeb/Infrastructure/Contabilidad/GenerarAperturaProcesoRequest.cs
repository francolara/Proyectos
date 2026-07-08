namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GenerarAperturaProcesoRequest
{
    public int IdEmpresa { get; init; }
    public short AnioApertura { get; init; }
    public byte MesSaldoHasta { get; init; }
    public decimal TipoCambioCompra { get; init; }
    public decimal TipoCambioVenta { get; init; }
    public string? UsuarioRegistro { get; init; }
}
