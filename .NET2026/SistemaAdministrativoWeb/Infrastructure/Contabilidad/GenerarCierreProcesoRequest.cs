namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GenerarCierreProcesoRequest
{
    public int IdEmpresa { get; init; }
    public short Anio { get; init; }
    public decimal TipoCambioCompra { get; init; }
    public decimal TipoCambioVenta { get; init; }
    public bool ProcesarGananciasPerdidas { get; init; }
    public bool ProcesarInventarios { get; init; }
    public string? UsuarioRegistro { get; init; }
}
