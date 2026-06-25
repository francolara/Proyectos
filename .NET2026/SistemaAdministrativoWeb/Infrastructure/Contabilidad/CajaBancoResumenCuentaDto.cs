namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CajaBancoResumenCuentaDto
{
    public decimal SaldoInicial { get; init; }
    public decimal IngresosMes { get; init; }
    public decimal EgresosMes { get; init; }
    public decimal SaldoFinal { get; init; }
}
