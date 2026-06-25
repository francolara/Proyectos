namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class GuardarBancoConfiguracionEmpresaRequest
{
    public int? IdBancoConfiguracionEmpresa { get; init; }
    public int IdEmpresa { get; init; }
    public int IdBanco { get; init; }
    public string NroCuentaCorriente { get; init; } = string.Empty;
    public string Titular { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public int IdPlanCuenta { get; init; }
    public bool Activo { get; init; }
    public string? UsuarioRegistro { get; init; }
}
