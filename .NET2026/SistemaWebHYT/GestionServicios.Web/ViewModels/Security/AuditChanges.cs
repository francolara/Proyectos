using System.Text.Json;
using System.Text.RegularExpressions;

namespace GestionServicios.Web.ViewModels.Security;

public sealed record AuditChange(string Field, string Before, string After);

public static partial class AuditChanges
{
    public static IReadOnlyList<AuditChange> From(JsonElement row)
    {
        if (!row.TryGetProperty("Cambios", out var changes) || changes.ValueKind != JsonValueKind.Object)
            return [];
        var before = new Dictionary<string, string>();
        var after = new Dictionary<string, string>();
        if (Find(changes, "Antes", "Anterior", "Before") is { } old) Flatten(old, "", before);
        if (Find(changes, "Despues", "Después", "Nuevo", "After") is { } current) Flatten(current, "", after);
        if (before.Count == 0 && after.Count == 0) Flatten(changes, "", after);
        return before.Keys.Union(after.Keys).Select(key =>
            new AuditChange(key, before.GetValueOrDefault(key, "—"), after.GetValueOrDefault(key, "—")))
            .Where(change => change.Before != change.After).ToArray();
    }

    private static JsonElement? Find(JsonElement element, params string[] names)
    {
        foreach (var property in element.EnumerateObject())
            if (names.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                return property.Value;
        return null;
    }

    private static void Flatten(JsonElement value, string field, Dictionary<string, string> result)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in value.EnumerateObject())
                    if (!Sensitive().IsMatch(property.Name))
                        Flatten(property.Value, Join(field, Label(property.Name)), result);
                break;
            case JsonValueKind.Array:
                var index = 0;
                foreach (var item in value.EnumerateArray())
                    Flatten(item, Join(field, $"#{++index}"), result);
                break;
            case JsonValueKind.String:
                var text = value.GetString() ?? "";
                if (text.TrimStart().StartsWith('{') || text.TrimStart().StartsWith('['))
                {
                    try { using var nested = JsonDocument.Parse(text); Flatten(nested.RootElement, field, result); }
                    catch (JsonException) { }
                }
                else if (field.StartsWith("Permisos", StringComparison.Ordinal) && text.Split('.') is [var module, _])
                    result[field] = SecurityDisplay.Module(module) + " · " + SecurityDisplay.Permission(text);
                else result[field] = text.Length == 0 ? "—" : text;
                break;
            default:
                result[field] = value.ValueKind switch
                {
                    JsonValueKind.True => "Sí", JsonValueKind.False => "No", JsonValueKind.Null => "—", _ => value.ToString()
                };
                break;
        }
    }

    private static string Join(string prefix, string name) => prefix.Length == 0 ? name : $"{prefix} / {name}";

    private static string Label(string code) => code switch
    {
        "Email" => "Correo",
        "Name" => "Nombre del perfil",
        "PerfilId" => "Perfil de acceso",
        "NumeroOCCliente" => "Número de OC del cliente",
        "NumeroActa" => "Número de guía",
        "NumeroInterno" => "Número interno",
        "FechaRegistro" => "Fecha de registro",
        _ => SecurityDisplay.Humanize(CamelCase().Replace(code, " "))
    };

    [GeneratedRegex("password|contrase|clave|hash|stamp|sello|token|cookie|secret|credencial|objetor2|storagekey", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex Sensitive();

    [GeneratedRegex("(?<=[a-z])(?=[A-Z])")]
    private static partial Regex CamelCase();
}
