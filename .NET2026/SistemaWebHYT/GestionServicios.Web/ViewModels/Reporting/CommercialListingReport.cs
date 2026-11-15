using System.Text.Json;
using GestionServicios.Web.ViewModels.Billing;
using static GestionServicios.Web.ViewModels.Reporting.ListingReport;

namespace GestionServicios.Web.ViewModels.Reporting;

internal static class CommercialListingReport
{
    public static ListingReport Documents(bool sales, IReadOnlyList<JsonElement> rows, IReadOnlyList<string> filters)
    {
        string[] headers = ["Comprobante", "Tipo", "Emisión", "Vencimiento", sales ? "Cliente" : "Proveedor",
            "Documento", "Forma de pago", "Moneda", "Importe", .. (sales
                ? new[] { "Detracción %", "Detracción", "Saldo pendiente", "Referencias", "SUNAT", "Factoring" }
                : new[] { "Pagado al proveedor", "Saldo al proveedor", "Detracción / pagada", "Estado de pago", "Trabajos vinculados" })];

        var table = new ListingTable(null, headers, rows.Select(row => DocumentRow(row, sales)));
        var totals = new ListingTable("Importes por moneda y tipo de comprobante", ["Moneda", "Tipo", "Registros", "Importe"],
            rows.GroupBy(row => (Currency: Text(row, "Moneda"), Type: Type(row)))
                .OrderBy(group => group.Key.Currency).ThenBy(group => group.Key.Type)
                .Select(group => new[] { group.Key.Currency, group.Key.Type, group.Count().ToString(), Sum(group, "Importe") }));
        var balances = rows.Where(row => Text(row, "Tipo") != "07").GroupBy(row => Text(row, "Moneda")).OrderBy(group => group.Key);
        if (sales)
            return new("Emisión de comprobantes", filters, rows.Count, table, totals,
                new ListingTable("Saldos de ventas por moneda", ["Moneda", "Detracción", "Saldo pendiente"],
                    balances.Select(group => new[] { group.Key, Sum(group, "DetraccionDocumento"), Money(group.Sum(CommercialPage.SalesBalance)) })));

        return new("Registro de compras", filters, rows.Count, table, totals,
            new ListingTable("Saldos a proveedores por moneda", ["Moneda", "Pagado", "Pendiente"],
                balances.Select(group => new[] { group.Key, Sum(group, "Pagado"), Money(group.Sum(PurchaseBalance)) })),
            new ListingTable("Detracciones de compras por moneda", ["Moneda", "Detracción", "Pagado", "Pendiente"],
                rows.Where(row => Text(row, "Tipo") != "07" && Value(row, "ImporteDetraccion") > 0)
                    .GroupBy(row => Text(row, "MonedaDetraccion")).OrderBy(group => group.Key)
                    .Select(group => new[] { group.Key, Sum(group, "ImporteDetraccion"), Sum(group, "Detraido"),
                        Money(group.Sum(row => Value(row, "ImporteDetraccion") - Value(row, "Detraido"))) })));
    }

    private static string[] DocumentRow(JsonElement row, bool sales)
    {
        string[] common = [CommercialPage.Number(row), Type(row), Date(row, "Fecha"), Date(row, "FechaVencimiento"),
            Text(row, "NombreTercero"), Text(row, "DocumentoTercero"), Text(row, "FormaPago"), Text(row, "Moneda"), Money(Amount(row, "Importe"))];
        var creditNote = Text(row, "Tipo") == "07";
        if (sales)
        {
            var lot = Text(row, "LoteId");
            var disbursement = Date(row, "FactoringFechaDesembolso");
            var rate = Value(row, "PorcentajeDetraccion");
            return [.. common, rate > 0 ? $"{rate:0.####}%" : "", Value(row, "DetraccionDocumento") > 0 ? Money(Amount(row, "DetraccionDocumento")) : "",
                Money(CommercialPage.SalesBalance(row)), References(row), Text(row, "EstadoSunat"),
                lot.Length == 0 ? "Sin factoring" : $"{CommercialPage.FactoringStatus(Text(row, "EstadoFactoring"))} · Lote {lot}" +
                    (disbursement.Length == 0 ? "" : $" · Fecha de desembolso {disbursement}")];
        }

        var paid = Value(row, "Pagado");
        var withholding = Value(row, "ImporteDetraccion");
        var withholdingPaid = Value(row, "Detraido");
        var balance = PurchaseBalance(row);
        return [.. common, creditNote ? "" : Money(paid), creditNote ? "" : Money(balance),
            withholding > 0 ? $"{Text(row, "MonedaDetraccion")} {Money(withholding)} / {Money(withholdingPaid)}" : "",
            creditNote ? "Nota de crédito" : balance <= 0 && withholdingPaid >= withholding ? "Pagado" :
                paid > 0 || withholdingPaid > 0 ? "Pago parcial" : "Pendiente", Text(row, "TrabajosVinculados")];
    }

    public static ListingReport Withholdings(IReadOnlyList<JsonElement> rows, IReadOnlyList<string> filters)
        => new("Pago de detracciones", filters, rows.Count,
            new ListingTable(null, ["Emisión", "Comprobante", "Proveedor", "RUC", "Código", "Moneda",
                "Detracción", "Nombre de la detracción", "%", "Pagado", "Saldo", "Último pago", "N.º de operación", "Estado"],
                rows.Select(row => new[] { Date(row, "Fecha"), CommercialPage.Number(row), Text(row, "NombreTercero"),
                    Text(row, "DocumentoTercero"), Text(row, "CodigoDetraccion"), Text(row, "MonedaDetraccion"),
                    Money(Amount(row, "ImporteDetraccion")), Text(row, "NombreDetraccion"), CommercialPage.WithholdingRate(row),
                    Money(Amount(row, "Pagado")), Money(Amount(row, "Saldo")),
                    Date(row, "FechaUltimoPago"), Text(row, "UltimaOperacion"),
                    Value(row, "Saldo") <= 0 ? "Pagada" : Value(row, "Pagado") > 0 ? "Pago parcial" : "Pendiente" })),
            new ListingTable("Totales por moneda", ["Moneda", "Detracción", "Pagado", "Saldo"],
                rows.GroupBy(row => Text(row, "MonedaDetraccion")).OrderBy(group => group.Key)
                    .Select(group => new[] { group.Key, Sum(group, "ImporteDetraccion"), Sum(group, "Pagado"), Sum(group, "Saldo") })));

    public static ListingReport Factoring(IReadOnlyList<JsonElement> rows, IReadOnlyList<string> filters)
        => new("Negociación de Factoring", filters, rows.Count,
            new ListingTable(null, ["Lote", "Financiera", "Negociación", "Fecha de desembolso", "Estado", "Referencia",
                "Porcentaje negociado", "Facturas", "Moneda", "Importe de facturas"],
                rows.Select(row => new[] { Text(row, "LoteId"), Text(row, "Financiera"), Date(row, "FechaNegociacion"),
                    Date(row, "FechaDesembolso"), CommercialPage.FactoringStatus(Text(row, "Estado")),
                    Text(row, "Referencia"), $"{Value(row, "Porcentaje"):0.####}%", Text(row, "Cantidad"),
                    Text(row, "Moneda"), Money(Amount(row, "Importe")) })),
            new ListingTable("Totales por moneda", ["Moneda", "Lotes", "Facturas", "Importe de facturas"],
                rows.GroupBy(row => Text(row, "Moneda")).OrderBy(group => group.Key)
                    .Select(group => new[] { group.Key, group.Count().ToString(),
                        group.Sum(row => Value(row, "Cantidad")).ToString("0"), Sum(group, "Importe") })));

    private static string Type(JsonElement row) => CommercialPage.TypeLabel(Text(row, "Tipo"));
    private static decimal Value(JsonElement row, string key) => Amount(row, key) ?? 0;
    private static string Sum(IEnumerable<JsonElement> rows, string key) => Money(rows.Sum(row => Value(row, key)));
    private static decimal PurchaseBalance(JsonElement row)
        => Value(row, "Importe") - Value(row, "DetraccionDocumento") - Value(row, "Pagado");

    private static string References(JsonElement row)
    {
        var values = new[] { ("OC", "NumeroOCCliente"), ("Confirmación", "NumeroConfirmacion"), ("Presupuestos", "PresupuestosVinculados") };
        return string.Join("\n", values.Select(item => (item.Item1, Value: Text(row, item.Item2)))
            .Where(item => item.Value.Length > 0).Select(item => $"{item.Item1} {item.Value}"));
    }
}
