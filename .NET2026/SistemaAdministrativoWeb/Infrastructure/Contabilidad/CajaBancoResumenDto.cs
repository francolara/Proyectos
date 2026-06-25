namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CajaBancoResumenDto
{
    public int IdMovimientoBanco { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public int NumeroMovimiento { get; init; }
    public int IdEmpresa { get; init; }
    public int IdBancoConfiguracionEmpresa { get; init; }
    public int IdBanco { get; init; }
    public string NroCuentaCorriente { get; init; } = string.Empty;
    public string Titular { get; init; } = string.Empty;
    public string CodigoBanco { get; init; } = string.Empty;
    public string NombreBanco { get; init; } = string.Empty;
    public string CodigoMoneda { get; init; } = string.Empty;
    public string NombreMoneda { get; init; } = string.Empty;
    public string TipoMovimiento { get; init; } = string.Empty;
    public string IdOpeBancaria { get; init; } = string.Empty;
    public string TipoOperacion { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public int? IdPersona { get; init; }
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public string NombrePersona { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string Glosa { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public decimal Ingreso { get; init; }
    public decimal Egreso { get; init; }
    public bool Activo { get; init; }
}
