namespace GestionServicios.Application.Budgets;

public sealed record BudgetLineAmounts(decimal Quantity, decimal? OriginPrice, string OriginCurrency, string OriginTax,
    decimal? UnitPrice, decimal? Net, decimal? Tax, decimal? Total);
public sealed record BudgetTotals(decimal? Net, decimal? Tax, decimal? Total, bool Complete);

public static class BudgetAmounts
{
    public const string Policy = "LINEA_2_V1";
    private const decimal Maximum = 99999999999999999.99m;
    private const decimal UnitMaximum = 9999999999999.99m;
    public static decimal Round(decimal value) => decimal.Round(value, 2, MidpointRounding.AwayFromZero);

    public static BudgetLineAmounts Calculate(decimal quantity, decimal? originPrice, string originCurrency,
        string originTax, string currency, string taxTreatment, decimal rate)
    {
        if (quantity <= 0 || quantity != Round(quantity) || quantity > UnitMaximum)
            throw new ArgumentException("La cantidad debe ser positiva y tener hasta dos decimales.");
        if (originPrice.HasValue && (originPrice < 0 || originPrice != Round(originPrice.Value) || originPrice > UnitMaximum))
            throw new ArgumentException("El precio debe ser positivo o cero y tener hasta dos decimales.");
        if (rate is < 0 or > 1 || originTax is not ("CON_IGV" or "SIN_IGV") || taxTreatment is not ("CON_IGV" or "SIN_IGV"))
            throw new ArgumentException("Revisa el tratamiento y la tasa de IGV.");
        if (currency is not ("PEN" or "USD") || originCurrency != currency)
            throw new ArgumentException("Revaloriza la partida en la moneda del presupuesto. No se realiza conversión cambiaria.");
        if (!originPrice.HasValue) return new(quantity, null, originCurrency, originTax, null, null, null, null);
        try
        {
            var netUnit = originTax == "CON_IGV" ? originPrice.Value / (1 + rate) : originPrice.Value;
            var unit = Round(taxTreatment == "CON_IGV" ? netUnit * (1 + rate) : netUnit);
            var amount = Round(checked(quantity * unit));
            var net = taxTreatment == "CON_IGV" ? Round(amount / (1 + rate)) : amount;
            var tax = taxTreatment == "CON_IGV" ? amount - net : Round(net * rate);
            var total = net + tax;
            if (unit > UnitMaximum || net > Maximum || tax > Maximum || total > Maximum) throw new OverflowException();
            return new(quantity, originPrice, originCurrency, originTax, unit, net, tax, total);
        }
        catch (OverflowException) { throw new ArgumentException("La cantidad o precio excede el importe máximo permitido."); }
    }

    public static BudgetTotals Sum(IReadOnlyList<BudgetLineAmounts> lines)
    {
        if (lines.Count == 0 || lines.Any(line => !line.Total.HasValue)) return new(null, null, null, false);
        try
        {
            var net = lines.Sum(line => line.Net!.Value);
            var tax = lines.Sum(line => line.Tax!.Value);
            var total = lines.Sum(line => line.Total!.Value);
            if (net > Maximum || tax > Maximum || total > Maximum) throw new OverflowException();
            return new(net, tax, total, true);
        }
        catch (OverflowException) { throw new ArgumentException("El total excede el importe máximo permitido."); }
    }
}
