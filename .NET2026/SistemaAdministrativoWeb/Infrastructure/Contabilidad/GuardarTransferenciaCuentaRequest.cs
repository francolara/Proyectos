namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarTransferenciaCuentaRequest
{
    public int IdEmpresa { get; init; }
    public int IdBancoConfiguracionEmpresaEmisor { get; init; }
    public int IdBancoConfiguracionEmpresaReceptor { get; init; }
    public string IdOpeBancariaEmisor { get; init; } = string.Empty;
    public string IdOpeBancariaReceptor { get; init; } = string.Empty;
    public DateOnly FechaEmisionEmisor { get; init; }
    public DateOnly FechaEmisionReceptor { get; init; }
    public decimal TipoCambioEmisor { get; init; }
    public decimal TipoCambioReceptor { get; init; }
    public string? NumeroOperacionEmisor { get; init; }
    public string? NumeroOperacionReceptor { get; init; }
    public decimal ImporteEmisor { get; init; }
    public decimal ImporteReceptor { get; init; }
    public string GlosaEmisor { get; init; } = string.Empty;
    public string GlosaReceptor { get; init; } = string.Empty;
    public string? ObservacionEmisor { get; init; }
    public string? ObservacionReceptor { get; init; }
    public string? UsuarioRegistro { get; init; }
}
