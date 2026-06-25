namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class CajaBancoMovimientoDto
{
    public int IdMovimientoBanco { get; init; }
    public int? IdAsiento { get; init; }
    public int? NumeroAsiento { get; init; }
    public int NumeroMovimiento { get; init; }
    public int IdEmpresa { get; init; }
    public int IdBancoConfiguracionEmpresa { get; init; }
    public string NroCuentaCorriente { get; init; } = string.Empty;
    public string CodigoBanco { get; init; } = string.Empty;
    public string NombreBanco { get; init; } = string.Empty;
    public string TitularCuentaCorriente { get; init; } = string.Empty;
    public int? IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public string NombreMoneda { get; init; } = string.Empty;
    public string TipoMovimiento { get; init; } = string.Empty;
    public string IdOpeBancaria { get; init; } = string.Empty;
    public DateOnly FechaEmision { get; init; }
    public decimal TipoCambio { get; init; }
    public int? IdPersona { get; init; }
    public string NumeroDocumentoPersona { get; init; } = string.Empty;
    public string NombrePersona { get; init; } = string.Empty;
    public string NumeroDocumento { get; init; } = string.Empty;
    public string Glosa { get; init; } = string.Empty;
    public string Observacion { get; init; } = string.Empty;
    public decimal ImporteTotal { get; init; }
    public bool Activo { get; init; }
    public List<CajaBancoDetalleDto> Detalles { get; init; } = [];
}
