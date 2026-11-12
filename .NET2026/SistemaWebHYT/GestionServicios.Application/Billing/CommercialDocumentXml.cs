using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace GestionServicios.Application.Billing;

public static class CommercialDocumentXml
{
    private static readonly XNamespace Basic = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
    private static readonly XNamespace Aggregate = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
    public const int MaximumSize = 2 * 1024 * 1024;

    public static async Task<CommercialDocument> ReadAsync(Stream input, bool sales, CancellationToken ct)
    {
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        {
            Async = true, DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = MaximumSize
        });
        var xml = await XDocument.LoadAsync(reader, LoadOptions.PreserveWhitespace, ct);
        var root = xml.Root ?? throw new FormatException("El XML está vacío.");
        var note = root.Name == XName.Get("CreditNote", "urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2");
        if (!note && root.Name != XName.Get("Invoice", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"))
            throw new FormatException("Carga un XML UBL de factura, boleta o nota de crédito.");
        var type = note ? "07" : Required(root, "InvoiceTypeCode");
        if (type is not ("01" or "03" or "07") || !note && type == "07")
            throw new FormatException("El tipo de comprobante no es compatible.");
        var number = Required(root, "ID").Split('-', 2, StringSplitOptions.TrimEntries);
        if (number.Length != 2 || number.Any(string.IsNullOrWhiteSpace))
            throw new FormatException("El XML no contiene una serie y un número válidos.");
        var date = DateOnly.ParseExact(Required(root, "IssueDate"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
        var currency = Required(root, "DocumentCurrencyCode");
        var supplier = Party(root, "AccountingSupplierParty");
        var customer = Party(root, "AccountingCustomerParty");
        var other = sales ? customer : supplier;
        var lines = root.Elements(Aggregate + (note ? "CreditNoteLine" : "InvoiceLine")).Select(line =>
        {
            var item = line.Element(Aggregate + "Item");
            var quantity = line.Element(Basic + (note ? "CreditedQuantity" : "InvoicedQuantity"));
            return new CommercialLine(Text(item?.Element(Aggregate + "SellersItemIdentification"), "ID"),
                Text(item, "Description") ?? Required(item, "Name"),
                (string?)quantity?.Attribute("unitCode") ?? "NIU", Money(quantity?.Value),
                Money(Required(line.Element(Aggregate + "Price"), "PriceAmount")), Money(Required(line, "LineExtensionAmount")));
        }).ToArray();
        var terms = root.Elements(Aggregate + "PaymentTerms").ToArray();
        var dueDate = ParseDueDate(Text(root, "DueDate"));
        if (dueDate is null)
            dueDate = terms.Where(term => Text(term, "ID") == "FormaPago")
                .Select(term => ParseDueDate(Text(term, "PaymentDueDate"))).Min();
        var form = terms.Where(term => Text(term, "ID") == "FormaPago").Select(term => Text(term, "PaymentMeansID"))
            .FirstOrDefault(value => value is "Contado" or "Credito");
        var withholdingTerms = terms.Where(term =>
            string.Equals(Text(term, "ID"), "Detraccion", StringComparison.OrdinalIgnoreCase));
        var salesInvoice = sales && type == "01";
        var withholding = salesInvoice
            ? withholdingTerms.OrderByDescending(WithholdingPercent).FirstOrDefault()
            : withholdingTerms.FirstOrDefault();
        var withheld = withholding?.Element(Basic + "Amount");
        var discrepancy = root.Element(Aggregate + "DiscrepancyResponse");
        var reference = root.Element(Aggregate + "BillingReference")?.Element(Aggregate + "InvoiceDocumentReference");
        var total = Money(Required(root.Element(Aggregate + "LegalMonetaryTotal"), "PayableAmount"));
        var withholdingPercent = salesInvoice ? WithholdingPercent(withholding) : null;
        var withholdingAmount = withheld is not null && (!sales || salesInvoice) ? Money(withheld.Value) : 0;
        var withholdingCurrency = (string?)withheld?.Attribute("currencyID") ?? currency;
        if (withheld is null && withholdingPercent is decimal declaredRate)
            withholdingAmount = WithholdingEquivalent(total, declaredRate);
        var equivalent = withholdingAmount;
        if (withholdingAmount > 0 && withholdingCurrency != currency)
        {
            var rate = withholdingPercent ?? WithholdingPercent(withholding);
            if (rate is null) throw new FormatException("La detracción usa otra moneda y falta el porcentaje en el XML.");
            equivalent = WithholdingEquivalent(total, rate.Value);
        }
        return new CommercialDocument(type, number[0], number[1], date, currency, other.Id, other.Name,
            form == "Credito" ? "Crédito" : form ?? "No indicada", total, lines,
            withholdingAmount, withholdingCurrency, Text(withholding, "PaymentMeansID"),
            Text(reference, "ID") ?? Text(discrepancy, "ReferenceID"), Text(discrepancy, "Description"),
            [], sales ? supplier.Id : customer.Id, xml.ToString(SaveOptions.DisableFormatting),
            DetraccionDocumento: equivalent, FechaVencimiento: dueDate, PorcentajeDetraccion: withholdingPercent,
            CodigoMotivoNota: note ? Text(discrepancy, "ResponseCode") : null);
    }

    private static decimal? WithholdingPercent(XElement? term)
    {
        if (Text(term, "PaymentPercent") is not { } value) return null;
        var rate = Money(value);
        return rate is >= 0 and <= 100 ? rate : throw new FormatException("El porcentaje de detracción del XML está fuera del rango de 0 a 100.");
    }

    private static decimal WithholdingEquivalent(decimal total, decimal rate)
        => decimal.Round(total * rate / 100m, 2, MidpointRounding.AwayFromZero);

    private static (string Id, string Name) Party(XElement root, string kind)
    {
        var account = root.Element(Aggregate + kind);
        var party = account?.Element(Aggregate + "Party");
        var legal = party?.Element(Aggregate + "PartyLegalEntity");
        var tax = party?.Element(Aggregate + "PartyTaxScheme");
        var id = Text(party?.Element(Aggregate + "PartyIdentification"), "ID") ?? Text(tax, "CompanyID")
            ?? Required(account, "CustomerAssignedAccountID");
        return (id, Text(legal, "RegistrationName") ?? Text(tax, "RegistrationName")
            ?? Text(party?.Element(Aggregate + "PartyName"), "Name") ?? "");
    }

    private static string? Text(XElement? node, string key)
        => node?.Element(Basic + key)?.Value.Trim() is { Length: > 0 } text ? text : null;

    private static string Required(XElement? node, string key)
        => Text(node, key) ?? throw new FormatException($"Falta {key} en el XML.");

    private static DateOnly? ParseDueDate(string? value)
        => value is null ? null : DateOnly.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static decimal Money(string? value)
        => decimal.TryParse(value, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var amount)
            ? amount : throw new FormatException("El XML contiene una cantidad o un importe inválido.");
}
