using System.Globalization;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public sealed record ListadoImpresionViewModel(
    string Negocio,
    string Titulo,
    string ArchivoCsv,
    DateOnly Desde,
    DateOnly Hasta,
    string? Busqueda,
    string? Documento,
    IReadOnlyList<string> Columnas,
    int PrimeraColumnaImporte,
    IReadOnlyList<string?[]> Filas,
    IReadOnlyList<string?[]> Totales)
{
    public static string Importe(decimal importe) => importe.ToString("F2", CultureInfo.GetCultureInfo("es-PE"));
}
