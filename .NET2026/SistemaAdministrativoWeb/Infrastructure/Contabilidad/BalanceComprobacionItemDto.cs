namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BalanceComprobacionItemDto
{
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string ColBalance { get; init; } = string.Empty;
    public byte GradoCuenta { get; init; }
    public decimal DebAnt { get; init; }
    public decimal HabAnt { get; init; }
    public decimal DebMes { get; init; }
    public decimal HabMes { get; init; }
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
}
