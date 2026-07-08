namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public static class PleEstadoRegistroCatalogo
{
    public const string Vigente = "1";
    public const string PeriodoAnterior = "6";
    public const string Omitido = "8";
    public const string Ajustado = "9";

    public static readonly IReadOnlySet<string> ValoresValidos = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        Vigente,
        PeriodoAnterior,
        Omitido,
        Ajustado
    };
}
