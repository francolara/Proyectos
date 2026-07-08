namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class AnalisisCuentaRequest
{
    public int IdEmpresa { get; set; }
    public string Periodo { get; set; } = string.Empty;
    public string? CuentaDesde { get; set; }
    public string? CuentaHasta { get; set; }
    public string? Auxiliar { get; set; }
    public string Moneda { get; set; } = "PEN";
    public string Estado { get; set; } = "T";
    public string Tipo { get; set; } = "0";
}
