namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarTipoCambioRequest
{
    public int? IdTipoCambio { get; init; }
    public int IdCuentaAdministradora { get; init; }
    public DateOnly Fecha { get; init; }
    public string IdMoneda { get; init; } = string.Empty;
    public decimal Compra { get; init; }
    public decimal Venta { get; init; }
    public decimal CompraSbs { get; init; }
    public decimal VentaSbs { get; init; }
    public string Fuente { get; init; } = string.Empty;
    public string? UsuarioRegistro { get; init; }
    public bool Estado { get; init; }
}
