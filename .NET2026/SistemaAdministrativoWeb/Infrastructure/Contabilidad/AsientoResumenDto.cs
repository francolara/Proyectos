namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public class AsientoResumenDto
{
    public int IdAsiento { get; init; }
    public int IdEmpresa { get; init; }
    public int IdOrigen { get; init; }
    public string CodigoOrigen { get; init; } = string.Empty;
    public string NombreOrigen { get; init; } = string.Empty;
    public bool PermiteRegistroManual { get; init; }
    public short Ejercicio { get; init; }
    public byte Mes { get; init; }
    public string Periodo { get; init; } = string.Empty;
    public int NumeroAsiento { get; init; }
    public DateOnly FechaAsiento { get; init; }
    public string Glosa { get; init; } = string.Empty;
    public int IdMoneda { get; init; }
    public string CodigoMoneda { get; init; } = string.Empty;
    public string NombreMoneda { get; init; } = string.Empty;
    public string SimboloMoneda { get; init; } = string.Empty;
    public decimal TipoCambio { get; init; }
    public decimal TotalDebe { get; init; }
    public decimal TotalHaber { get; init; }
    public string Estado { get; init; } = string.Empty;
    public string? ReferenciaExterna { get; init; }
    public string? Observacion { get; init; }
}
