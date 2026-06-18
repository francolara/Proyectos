namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed class PaginacionViewModel
{
    public int PaginaActual { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas => TotalRegistros <= 0
        ? 1
        : (int)Math.Ceiling(TotalRegistros / (double)TamanoPagina);
    public bool TienePaginaAnterior => PaginaActual > 1;
    public bool TienePaginaSiguiente => PaginaActual < TotalPaginas;
}
