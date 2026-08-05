namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public static class PleLibroElectronicoCatalogo
{
    public const string LibroDiario51 = "5.1";
    public const string LibroDiario52 = "5.2";
    public const string PlanContable53 = "5.3";
    public const string PlanContable54 = "5.4";
    public const string LibroMayor61 = "6.1";

    public static readonly IReadOnlyDictionary<string, string> Nombres = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        [LibroDiario51] = "Libro Diario - Formato 5.1",
        [LibroDiario52] = "Libro Diario Simplificado - Formato 5.2",
        [LibroMayor61] = "Libro Mayor - Formato 6.1"
    };

    public static readonly IReadOnlyDictionary<string, string> CodigosSunat = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        [LibroDiario51] = "050100",
        [LibroDiario52] = "050200",
        [PlanContable53] = "050300",
        [PlanContable54] = "050400",
        [LibroMayor61] = "060100"
    };

    public static bool EsValido(string? codigo)
    {
        return !string.IsNullOrWhiteSpace(codigo) && Nombres.ContainsKey(codigo.Trim());
    }

    public static string Normalizar(string? codigo)
    {
        return codigo?.Trim() switch
        {
            LibroDiario52 => LibroDiario52,
            LibroMayor61 => LibroMayor61,
            _ => LibroDiario51
        };
    }

    public static string ObtenerNombre(string? codigo)
    {
        var normalizado = Normalizar(codigo);
        return Nombres.TryGetValue(normalizado, out var nombre)
            ? nombre
            : Nombres[LibroDiario51];
    }

    public static string ObtenerCodigoSunat(string? codigo)
    {
        var codigoTrabajo = codigo?.Trim() ?? string.Empty;
        return CodigosSunat.TryGetValue(codigoTrabajo, out var valor)
            ? valor
            : CodigosSunat[LibroDiario51];
    }

    public static string? ObtenerPlanComplementario(string? codigo)
    {
        return Normalizar(codigo) switch
        {
            LibroDiario51 => PlanContable53,
            LibroDiario52 => PlanContable54,
            _ => null
        };
    }
}
