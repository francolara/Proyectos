namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAplicacionNotaCreditoRequest
{
    public int IdEmpresa { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdPersona { get; init; }
    public DateOnly FechaAplicacion { get; init; }
    public decimal TipoCambio { get; init; }
    public int IdRegistroComprobante { get; init; }
    public int IdRegistroNotaCredito { get; init; }
    public decimal ImporteAplicado { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
}
