namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class TransferenciaCuentaResultadoDto
{
    public Guid IdTransferenciaCuenta { get; init; }
    public int IdMovimientoBancoEmisor { get; init; }
    public int NumeroMovimientoEmisor { get; init; }
    public int? NumeroAsientoEmisor { get; init; }
    public int IdMovimientoBancoReceptor { get; init; }
    public int NumeroMovimientoReceptor { get; init; }
    public int? NumeroAsientoReceptor { get; init; }
    public decimal ImporteEmisor { get; init; }
    public decimal ImporteReceptor { get; init; }
}
