using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;

namespace GestionServicios.Web.ViewModels.Reporting;

public sealed record ListingReport(string Title, IReadOnlyList<string> Filters, int Count, params ListingTable[] Tables)
{
    private static readonly CultureInfo Peru = CultureInfo.GetCultureInfo("es-PE");
    public static string Money(decimal? value) => value?.ToString("N2", Peru) ?? "—";
    public static string Date(DateOnly? date) => date?.ToString("dd/MM/yyyy", Peru) ?? "—";
    public static string Date(DateTime? date) => date is DateTime value ? Date(DateOnly.FromDateTime(value)) : "—";
    public static string Date(JsonElement row, string key) =>
        DateOnly.TryParseExact(Text(row, key), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)
            ? Date(date) : Text(row, key);
    public static string Text(JsonElement row, string key) => row.ValueKind == JsonValueKind.Object
        ? BudgetRecord.Text(row, key) : "";
    public static decimal? Amount(JsonElement row, string key) => row.ValueKind == JsonValueKind.Object && row.TryGetProperty(key, out var value) &&
        value.ValueKind == JsonValueKind.Number && value.TryGetDecimal(out var number) ? number : null;
    public static string Filter(string label, string? value) => $"{label}: {(string.IsNullOrWhiteSpace(value) ? "Todos" : value)}";
    public static string Range(DateOnly? from, DateOnly? to) => $"Desde {Date(from)} hasta {Date(to)}";

    public static async Task<IReadOnlyList<T>> AllAsync<T>(
        Func<int, Task<(IReadOnlyList<T> Rows, long Total)>> load, Func<T, long> id)
    {
        var records = new List<T>();
        var seen = new HashSet<long>();
        long expected = 0;
        for (var page = 1; ; page++)
        {
            var result = await load(page);
            if (page == 1) expected = result.Total;
            if (result.Total != expected || result.Rows.Any(row => !seen.Add(id(row))) ||
                (result.Rows.Count == 0 && records.Count < expected))
                throw new MaintenanceException("El listado cambió mientras se preparaba el reporte. Vuelve a generarlo.");

            records.AddRange(result.Rows);
            if (records.Count == expected) return records;
            if (records.Count > expected)
                throw new MaintenanceException("El listado cambió mientras se preparaba el reporte. Vuelve a generarlo.");
        }
    }
}

public sealed record ListingTable(string? Title, string[] Headers, IEnumerable<string[]> Rows);
