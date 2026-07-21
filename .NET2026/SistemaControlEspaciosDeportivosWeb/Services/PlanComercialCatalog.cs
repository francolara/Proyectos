namespace SistemaControlEspaciosDeportivosWeb.Services;

public sealed record PlanComercialDefinition(
    string Codigo,
    string Nombre,
    decimal PrecioMensual,
    int DiasPrueba,
    string TipoPlanInterno,
    int SedesPermitidas,
    int EspaciosPermitidos,
    int UsuariosPermitidos);

public static class PlanComercialCatalog
{
    public const string Prueba = "PRUEBA";
    public const string Esencial = "ESENCIAL";
    public const string Pro = "PRO";

    private static readonly IReadOnlyDictionary<string, PlanComercialDefinition> Planes =
        new Dictionary<string, PlanComercialDefinition>(StringComparer.OrdinalIgnoreCase)
        {
            [Prueba] = new(Prueba, "Prueba", 0m, 15, "Basico", 1, 1, 1),
            [Esencial] = new(Esencial, "Esencial", 49.90m, 0, "Basico", 1, 5, 2),
            [Pro] = new(Pro, "Pro", 99.90m, 0, "Full", 3, 12, 3)
        };

    public static string Normalizar(string? codigo) => (codigo ?? string.Empty).Trim().ToUpperInvariant() switch
    {
        Prueba or "GRATIS" => Prueba,
        Esencial or "EMPRENDEDOR" => Esencial,
        Pro or "PROFESIONAL" => Pro,
        _ => Prueba
    };

    public static PlanComercialDefinition Obtener(string? codigo) => Planes[Normalizar(codigo)];
}
