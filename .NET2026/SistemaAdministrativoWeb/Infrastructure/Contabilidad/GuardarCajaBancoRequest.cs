namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarCajaBancoRequest
{
    public int? IdMovimientoBanco { get; init; }
    public int IdEmpresa { get; init; }
    public int IdBancoConfiguracionEmpresa { get; init; }
    public string TipoMovimiento { get; init; } = string.Empty;
    public string IdOpeBancaria { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public decimal TipoCambio { get; init; }
    public int? IdPersona { get; init; }
    public string? NumeroDocumento { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public string? Observacion { get; init; }
    public decimal ImporteTotal { get; init; }
    public string? UsuarioRegistro { get; init; }
    public List<GuardarCajaBancoDetalleRequest> Detalles { get; init; } = [];
}
