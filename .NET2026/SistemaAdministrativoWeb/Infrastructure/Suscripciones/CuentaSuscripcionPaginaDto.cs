namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class CuentaSuscripcionPaginaDto
{
    public int PaginaActual { get; init; } = 1;
    public int TamanoPagina { get; init; } = 10;
    public int TotalFiltrado { get; init; }
    public int TotalPaginas { get; init; } = 1;
    public int TotalCuentas { get; init; }
    public int CuentasActivas { get; init; }
    public int CuentasEnPrueba { get; init; }
    public int CuentasSuspendidasOBaja { get; init; }
    public int CobrosRegistrados { get; init; }
    public int CobrosPendientesAplicacion { get; init; }
    public decimal MontoCobradoMes { get; init; }
    public IReadOnlyCollection<CuentaSuscripcionResumenDto> Cuentas { get; init; } = [];
}
