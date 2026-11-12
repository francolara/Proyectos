using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace GestionServicios.Application.Payroll;

public sealed record TelecreditAccount(string TipoCuenta, string NumeroCuenta, string Moneda, byte[]? VersionFila = null);
public sealed record TelecreditBatch(long PlanId, TelecreditAccount? Cargo, TelecreditPayment[]? Abonos);
public sealed record TelecreditPayment(long TrabajadorId, string Nombre, string? TipoDocumento, string Documento,
    string? Banco, string? Cuenta, string? Cci, string? TipoCuenta, string? Moneda, decimal Neto, bool TieneMovimientos);
public sealed record TelecreditFile(string Name, byte[] Bytes);
public sealed record TelecreditResult(TelecreditFile? File, IReadOnlyList<string> Errors);

public interface ITelecreditRepository
{
    Task<TelecreditAccount?> GetChargeAccountAsync(string actor, CancellationToken ct);
    Task<long> SaveChargeAccountAsync(string actor, TelecreditAccount account, CancellationToken ct);
    Task<TelecreditBatch> ReadBatchAsync(string actor, long planId, byte[] version, string kind, CancellationToken ct);
}

public static partial class Telecredit
{
    private const decimal MaxAmount = 99999999999999.99m;
    private const long MaxChecksum = 999999999999999L;

    [GeneratedRegex(@"^[A-Z .()#/@&-]+$")]
    private static partial Regex AllowedText();

    public static string NormalizeAccount(string? value) => value?.Trim().Replace(" ", "").Replace("-", "") ?? "";

    public static bool ValidIdentity(string? type, string? number)
        => number is { Length: >= 1 and <= 12 } && number.All(char.IsAsciiLetterOrDigit) &&
            (type is "3" or "4" || type == "1" && Digits(number, 8));

    public static bool HasChargeAccountFormat(string account) => Digits(account, 13);

    public static string? ChargeAccountCurrency(string account)
        => HasChargeAccountFormat(account) ? account[10] switch { '0' => "PEN", '1' => "USD", _ => null } : null;

    public static bool IsBcp(string? bank) => NormalizeText(bank) is
        "BCP" or "BANCO DE CREDITO" or "BANCO DE CREDITO DEL PERU" or "BANCO DE CREDITO DEL PERU (BCP)";

    public static TelecreditResult Generate(TelecreditBatch batch, string kind, DateOnly date, DateOnly today,
        string? reference, bool validateIdentity)
    {
        List<string> errors = [];
        var subtype = kind switch { "PLANILLA" => "X", "EXTERNO" => "4", _ => "" };
        var title = subtype == "X" ? "PAGO DE SUELDOS" : "PAGO HONORARIOS";
        var description = NormalizeText(reference);
        if (subtype.Length == 0) errors.Add("Selecciona Planilla u Honorarios.");
        if (date < today) errors.Add("La fecha de proceso debe ser igual o posterior a hoy en Perú.");
        if (!ValidText(description, 40)) errors.Add("Indica una referencia de lote de hasta 40 caracteres admitidos por BCP, sin números.");

        var cargo = batch.Cargo;
        if (cargo is null || cargo.TipoCuenta is not ("C" or "M") ||
            ChargeAccountCurrency(cargo.NumeroCuenta) != cargo.Moneda)
            errors.Add("Configura una cuenta de cargo BCP válida en Configuración HYT.");
        else if (cargo.Moneda != "PEN")
            errors.Add("La cuenta de cargo está en dólares. Para exportar esta planilla en soles, configura una cuenta de cargo en soles.");

        var payments = batch.Abonos ?? [];
        if (payments.Length == 0) errors.Add("No hay abonos del tipo seleccionado en la planificación pendiente.");
        if (payments.Length > 999999) errors.Add("El lote supera la cantidad de abonos admitida por BCP.");
        if (errors.Count > 0 || cargo is null) return new(null, errors);

        var details = new List<string>(payments.Length);
        decimal total = 0;
        long checksum = AccountValue(cargo.NumeroCuenta, 3);
        foreach (var payment in payments)
        {
            var name = NormalizeText(payment.Nombre);
            var document = payment.Documento.Trim().ToUpperInvariant();
            var bankIsBcp = IsBcp(payment.Banco);
            var type = bankIsBcp ? payment.TipoCuenta : "B";
            var account = NormalizeAccount(bankIsBcp ? payment.Cuenta : payment.Cci);
            var issues = new List<string>();
            if (payment.TieneMovimientos) issues.Add("ya tiene movimientos de pago o planilla en el período");
            if (!ValidText(name, 75)) issues.Add("el nombre debe tener hasta 75 caracteres admitidos por BCP");
            if (!ValidIdentity(payment.TipoDocumento, document)) issues.Add("tipo o número de documento de identidad inválido");
            if (string.IsNullOrWhiteSpace(payment.Banco)) issues.Add("falta el banco de abono");
            if (payment.Moneda != cargo.Moneda) issues.Add("la cuenta de abono debe estar configurada en soles");
            if (DestinationAccountError(type, bankIsBcp, account) is { } accountError) issues.Add(accountError);
            if (account == cargo.NumeroCuenta)
                issues.Add("la cuenta de abono coincide con la cuenta de cargo");
            if (payment.Neto <= 0 || payment.Neto > MaxAmount || decimal.Round(payment.Neto, 2) != payment.Neto)
                issues.Add("el neto a abonar debe ser positivo y caber en el campo de importe con dos decimales");

            if (issues.Count != 0)
            {
                errors.Add($"{payment.Nombre} (trabajador {payment.TrabajadorId}) — {string.Join(", ", issues)}.");
                continue;
            }

            total += payment.Neto;
            checksum += AccountValue(account, type == "B" ? 10 : 3);
            details.Add("2" + type + account.PadRight(20) + payment.TipoDocumento + document.PadRight(12)
                + "   " + name.PadRight(75) + title.PadRight(40) + title.PadRight(20) + "0001"
                + FormatAmount(payment.Neto) + (validateIdentity ? "S" : "N"));
        }

        if (total > MaxAmount) errors.Add("El total del lote supera el ancho del importe de cabecera.");
        if (checksum > MaxChecksum) errors.Add("El total de control supera las 15 posiciones admitidas por BCP.");
        if (errors.Count > 0) return new(null, errors);

        var processDate = date.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
        var header = "1" + details.Count.ToString("D6", CultureInfo.InvariantCulture) + processDate + subtype
            + cargo.TipoCuenta + "0001" + cargo.NumeroCuenta.PadRight(20) + FormatAmount(total)
            + description.PadRight(40) + checksum.ToString("D15", CultureInfo.InvariantCulture);
        if (header.Length != 113 || details.Any(detail => detail.Length != 195))
            throw new InvalidOperationException("El archivo de Haberes tiene una longitud inesperada.");

        var category = subtype == "X" ? "5TA" : "4TA";
        var contents = header + "\r\n" + string.Join("\r\n", details) + "\r\n";
        return new(new($"BCP_HABERES_{category}_{processDate}.txt", Encoding.ASCII.GetBytes(contents)), []);
    }

    private static string? DestinationAccountError(string? type, bool bankIsBcp, string number)
    {
        if (!bankIsBcp)
            return Digits(number, 20) ? null : "registra el CCI de 20 dígitos para el abono interbancario";

        var length = type switch { "A" => 14, "C" or "M" => 13, _ => 0 };
        if (length == 0) return "selecciona el tipo de cuenta BCP en la ficha del trabajador";
        var label = type == "A" ? "de ahorros" : type == "C" ? "corriente" : "maestra";
        if (!Digits(number, length)) return $"el número de cuenta BCP {label} debe tener {length} dígitos";

        var position = length - 2;
        return number[^3] switch
        {
            '0' => null,
            '1' => $"el dígito {position} de la cuenta BCP {label} indica dólares (1). Registra una cuenta en soles para este pago",
            var digit => $"el dígito {position} de la cuenta BCP {label} es {digit}. BCP admite 0 para soles o 1 para dólares. Verifica el número en la ficha del trabajador"
        };
    }

    private static bool Digits(string text, int length) => text.Length == length && text.All(char.IsAsciiDigit);
    private static long AccountValue(string number, int skip) => long.Parse(number.AsSpan(skip), CultureInfo.InvariantCulture);
    private static string FormatAmount(decimal value) => value.ToString("00000000000000.00", CultureInfo.InvariantCulture);
    private static bool ValidText(string text, int length) => text.Length <= length && AllowedText().IsMatch(text);

    private static string NormalizeText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        return string.Concat(value.Trim().ToUpperInvariant().Normalize(NormalizationForm.FormD)
            .Where(character => CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark));
    }
}
