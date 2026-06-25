namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarAplicacionNotaCreditoResultDto
{
    public int IdAplicacionNotaCredito { get; init; }
    public int IdEmpresa { get; init; }
    public string ModuloOperacion { get; init; } = string.Empty;
    public int IdPersona { get; init; }
    public DateOnly FechaAplicacion { get; init; }
    public int IdRegistroComprobante { get; init; }
    public int IdRegistroNotaCredito { get; init; }
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal ImporteAplicado { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public string Glosa { get; init; } = string.Empty;
}
