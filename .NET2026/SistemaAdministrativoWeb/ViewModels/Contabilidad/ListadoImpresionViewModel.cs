namespace SistemaAdministrativoWeb.ViewModels.Contabilidad;

public sealed record ListadoImpresionViewModel(
    string Titulo,
    string EmpresaNombre,
    IReadOnlyList<(string Etiqueta, string? Valor)> Filtros,
    IReadOnlyList<string> Columnas,
    IReadOnlyList<string[]> Filas,
    IReadOnlyList<string[]>? Totales = null,
    int PrimeraColumnaImporte = int.MaxValue,
    bool Horizontal = false);
