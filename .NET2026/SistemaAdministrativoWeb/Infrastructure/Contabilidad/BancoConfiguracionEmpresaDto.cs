namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class BancoConfiguracionEmpresaDto
{
    public int IdBancoConfiguracionEmpresa { get; init; }
    public int IdEmpresa { get; init; }
    public int IdBanco { get; init; }
    public string CodigoBanco { get; init; } = string.Empty;
    public string NombreBanco { get; init; } = string.Empty;
    public string NroCuentaCorriente { get; init; } = string.Empty;
    public string Titular { get; init; } = string.Empty;
    public int? IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public string NombreMoneda { get; init; } = string.Empty;
    public int IdPlanCuenta { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public bool Activo { get; init; }
    public DateTime FechaRegistro { get; init; }
    public string UsuarioRegistro { get; init; } = string.Empty;
}
