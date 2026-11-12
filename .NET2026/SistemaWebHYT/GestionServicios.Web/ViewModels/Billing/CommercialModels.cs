using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Budgets;

namespace GestionServicios.Web.ViewModels.Billing;

public sealed record CommercialPage(string Scope, JsonElement Data, bool CanEdit, DateOnly Today,
    DateOnly? From = null, DateOnly? To = null, string? Type = null, string? Search = null, int Page = 1, bool WithoutFactoring = false, bool? HasFactoring = null)
{
    public bool Sales => Scope == "ventas";
    public string Title => Sales ? "Emisión de comprobantes" : "Registro de compras";
    public static string Text(JsonElement row, string key) => row.ValueKind == JsonValueKind.Object ? BudgetRecord.Text(row, key) : "";
    public static decimal Amount(JsonElement row, string key)
        => row.TryGetProperty(key, out var value) && value.TryGetDecimal(out var number) ? number : 0;
    public static string Money(decimal value) => value.ToString("N2", CultureInfo.GetCultureInfo("es-PE"));
    public static string Date(string value) => DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture,
        DateTimeStyles.None, out var date) ? date.ToString("dd/MM/yyyy") : value;
    public static string TypeLabel(string value) => value switch { "01" => "Factura", "03" => "Boleta", "07" => "Nota de crédito", _ => value };
    public static string Number(JsonElement row) => Text(row, "Serie") is { Length: > 0 } series ? $"{series}-{Text(row,"Numero")}" : "Borrador";
}

public sealed record CommercialForm(string Scope, JsonElement Data, CommercialInput Input, long? ClientId, long? BudgetId, DateOnly Today)
{
    public bool Sales => Scope == "ventas";
    public decimal? TaxRate => ReadTaxRate(Data);

    public static decimal? ReadTaxRate(JsonElement data)
        => data.TryGetProperty("TasaIgv", out var value) && value.ValueKind == JsonValueKind.Number
            ? value.GetDecimal() : null;
}

// Firma: FRANCO LARA - 11/11/2026 | Valida el plazo de crédito y calcula el vencimiento de los comprobantes manuales de compras y ventas y exige cantidades enteras en el detalle y conserva el motivo SUNAT de la nota de crédito.
public sealed class CommercialInput
{
    public static IReadOnlyList<int> CreditTerms { get; } = [15, 30, 45, 60, 90, 120];

    public Guid OperacionId { get; set; } = Guid.NewGuid();
    public string Tipo { get; set; } = "01";
    public string? Serie { get; set; }
    public string? Numero { get; set; }
    public DateOnly Fecha { get; set; }
    public string Moneda { get; set; } = "PEN";
    public string? DocumentoTercero { get; set; }
    public string? NombreTercero { get; set; }
    public string FormaPago { get; set; } = "Contado";
    public int? PlazoPagoDias { get; set; }
    public long? OrdenCompraId { get; set; }
    public string? CodigoDetraccion { get; set; }
    public string? MonedaDetraccion { get; set; }
    public decimal? ImporteDetraccion { get; set; }
    public string? ReferenciaNota { get; set; }
    public string? MotivoNota { get; set; }
    public string? CodigoMotivoNota { get; set; }
    public List<CommercialLine> Detalle { get; set; } = [];
    public List<long> Presupuestos { get; set; } = [];

    public DateOnly? DueDate => Fecha == default ? null : FormaPago switch
    {
        "Contado" => Fecha,
        "Crédito" when PlazoPagoDias is int days && CreditTerms.Contains(days)
            && Fecha.DayNumber <= DateOnly.MaxValue.DayNumber - days => Fecha.AddDays(days),
        _ => null
    };

    public CommercialDocument ToDocument(decimal rate, bool sales)
    {
        var dueDate = DueDate;
        if (dueDate is null)
            throw new FormatException("Selecciona Contado o Crédito con un plazo que permita calcular una fecha de vencimiento válida.");
        if (Detalle.Any(line => line.Cantidad != decimal.Truncate(line.Cantidad)))
            throw new FormatException("La cantidad de cada ítem debe ser un número entero.");

        var lines = Detalle.Select(line => line with { Importe = BudgetAmounts.Round(line.Cantidad * line.Precio) }).ToArray();
        var subtotal = lines.Sum(line => line.Importe);
        var total = subtotal + BudgetAmounts.Round(subtotal * rate);
        var withholdingCode = !sales || Tipo == "01" ? CodigoDetraccion?.Trim() : null;
        var withholdingCurrency = sales || string.IsNullOrEmpty(withholdingCode) || string.IsNullOrEmpty(MonedaDetraccion)
            ? Moneda : MonedaDetraccion;
        var foreignDeposit = withholdingCurrency != Moneda ? ImporteDetraccion.GetValueOrDefault() : 0;
        var isCreditNote = Tipo == "07";
        if (sales && isCreditNote && (string.IsNullOrWhiteSpace(ReferenciaNota) ||
            CodigoMotivoNota is not { Length: 2 } || !CodigoMotivoNota.All(char.IsAsciiDigit)))
            throw new FormatException("Selecciona la factura y el motivo SUNAT de la nota de crédito.");

        return new(Tipo, Serie?.Trim().ToUpperInvariant(), Numero?.Trim(), Fecha, Moneda,
            DocumentoTercero?.Trim() ?? "", NombreTercero?.Trim() ?? "", FormaPago, total, lines, foreignDeposit,
            withholdingCurrency, withholdingCode, isCreditNote ? ReferenciaNota?.Trim() : null,
            isCreditNote && !sales ? MotivoNota?.Trim() : null, Presupuestos,
            OperacionId: OperacionId, FechaVencimiento: dueDate, OrdenCompraId: sales ? OrdenCompraId : null,
            CodigoMotivoNota: isCreditNote && sales ? CodigoMotivoNota : null);
    }
}
