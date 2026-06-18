namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AsientoPreviewLineaDto
{
    public short Item { get; init; }
    public string ComponenteContable { get; init; } = string.Empty;
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string NaturalezaMovimiento { get; init; } = string.Empty;
    public decimal Debe { get; init; }
    public decimal Haber { get; init; }
    public string GlosaDetalle { get; init; } = string.Empty;
}
