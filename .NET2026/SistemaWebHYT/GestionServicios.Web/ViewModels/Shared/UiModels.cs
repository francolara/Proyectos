using GestionServicios.Contracts;
namespace GestionServicios.Web.ViewModels.Shared;

public sealed record DatePickerModel(string Name, string Label, DateOnly? Value = null, bool ReadOnly = false, bool Required = false, DateOnly? Min = null, DateOnly? Max = null)
{
    public string Id => "date_" + System.Text.RegularExpressions.Regex.Replace(Name, "[^a-zA-Z0-9_]", "_");
}
public sealed record FilterChoice(string Value,string Label);
public sealed record PayrollCatalogPicker(string Id, string TargetId, string Label,
    IReadOnlyList<GestionServicios.Application.Payroll.PayrollCatalogEntry> Entries, string? RegimeFieldId = null);
public sealed record FilterSelect(string Name,string Label,string Value,IReadOnlyList<FilterChoice> Choices);
public sealed record FilterCheck(string Name,string Label,bool Checked);
public sealed record StateKpi(string Code,string Label,long Count);
public sealed record FilterText(string Name,string Label,string Value);
public sealed class FilterModel
{
    public IReadOnlyList<FilterCheck> Checks {get;init;}=[];
    public IReadOnlyList<FilterText> Texts {get;init;}=[];
    public IReadOnlyList<FilterSelect> Selects {get;init;}=[];
    public string Search { get; init; } = "";
    public DateOnly? From { get; init; }
    public DateOnly? To { get; init; }
    public bool WithDates { get; init; }
    public string Placeholder { get; init; } = "Código o descripción";
    public string Controller { get; init; } = "Home";
    public string Action { get; init; } = "Index";
    public static FilterModel DocumentDefaults(string controller)
    {
        var today = DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(DateTimeOffset.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
        return new() { Controller = controller, WithDates = true, From = today.AddMonths(-1), To = today };
    }
}
public sealed record GridAction(string Label,string Url,bool Popup=false,bool HasRecord=false);
public sealed record GridCellAction(string Url,bool DoubleClick=false,string? Label=null,bool Navigate=false);
public sealed record GridRow(string Id, IReadOnlyList<string> Cells, bool CanView = true, bool CanEdit = false,IReadOnlyList<GridAction>? Actions=null,IReadOnlyDictionary<int,GridCellAction>? CellActions=null);
public sealed class GridModel
{
    public string Label { get; init; } = "Resultados";
    public string Controller { get; init; } = "Home";
    public string Action { get; init; } = "Index";
    public IReadOnlyList<string> Headers { get; init; } = [];
    public IReadOnlyList<GridRow> Rows { get; init; } = [];
    public long TotalRecords { get; init; }
    public int Page { get; init; } = 1;
    public long TotalPages => TotalRecords / 20 + (TotalRecords % 20 == 0 ? 0 : 1);
    public string? Error { get; init; }
    public Dictionary<string,string> Query { get; init; } = [];
    public Dictionary<string,string> RowQuery { get; init; } = [];
    public Dictionary<string,string> QueryFor(int page) => new(Query) { ["page"] = page.ToString(System.Globalization.CultureInfo.InvariantCulture) };
    public void Validate()
    {
        if (Rows.Count > 20 || Page < 1 || TotalRecords < 0 || Rows.Any(row => row.Cells.Count != Headers.Count))
            throw new ArgumentException("La grilla requiere una página SQL válida de hasta 20 registros y columnas coherentes.");
    }
}
public sealed class RecordFormModel
{
    public string Title { get; init; } = "Registro";
    public string Section { get; init; } = "";
    public string? Id { get; init; }
    public bool ReadOnly { get; init; }
    public bool IsNew { get; init; }
    public string Controller { get; init; } = "Home";
    public string SaveAction { get; init; } = "Save";
    public string? FieldsPartial { get; init; }
    public object? Fields { get; init; }
    public string? Error { get; init; }
    public RecordTrace? Trace { get; init; }
}
public sealed record ModalModel(string Id, string Title, string? Subtitle = null, string? ContentPartial = null, object? Content = null);
public static class PeruDate
{
    public static string Display(DateTimeOffset value) => TimeZoneInfo.ConvertTime(value, TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).ToString("dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.GetCultureInfo("es-PE"));
}
