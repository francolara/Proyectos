using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace GestionServicios.Application.Maintenance;

// Firma: FRANCO LARA - 05/11/2026 | Códigos del servidor, tarifas contextuales e identidad por persona sin tipo de documento y lista de enlaces HTTP/S con descripción. Validación de ubigeo y dirección.
public sealed record ValidatedInput(string Payload, byte[]? Version, IReadOnlyList<ValidationIssue> Issues);

public static partial class MaintenanceInput
{
    [GeneratedRegex(@"^\d{1,13}(\.\d{1,2})?$")]
    private static partial Regex PricePattern();

    public static ValidatedInput Validate(EntityDefinition entity, IReadOnlyDictionary<string, string?> fields,
        long? id, string? rowVersion, IEnumerable<int> zones, IEnumerable<ClientDocumentLink>? documentLinks = null)
    {
        var issues = new List<ValidationIssue>();
        var values = new Dictionary<string, object?>();
        byte[]? version = null;
        if (id.HasValue)
        {
            try { version = Convert.FromBase64String(rowVersion ?? ""); }
            catch (FormatException) { }
            if (id <= 0 || version?.Length != 8)
                issues.Add(new("", "Vuelve a abrir el registro para editarlo; falta su versión original."));
        }
        foreach (var field in entity.Fields.Where(field => field.Kind is not ("ReadOnly" or "AutoCode" or "Hidden")))
        {
            var text = fields.GetValueOrDefault(field.Name)?.Trim();
            if(entity.Key=="clientes")
            {
                var entityType=fields.GetValueOrDefault("TipoEntidad");
                if(entityType=="JURIDICA" && field.Name is "Nombres" or "Apellidos" || entityType=="NATURAL" && field.Name=="RazonSocial")
                {values[field.Name]=null;continue;}
            }
            if (string.IsNullOrEmpty(text))
            {
                values[field.Name] = null;
                if (field.Required) issues.Add(new(field.Name, $"Completa {field.Label.ToLowerInvariant()}."));
                continue;
            }
            if (field.MaxLength > 0 && text.Length > field.MaxLength)
                issues.Add(new(field.Name, $"El máximo es {field.MaxLength} caracteres."));
            object? value = text;
            var valid = true;
            switch (field.Kind)
            {
                case "Boolean":
                    valid = bool.TryParse(text, out var flag);
                    value = flag;
                    break;
                case "Reference":
                    valid = long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var reference) && reference > 0;
                    if (field.Lookup is "unidades" or "zonas" or "especialidades" or "tipos-servicio") valid &= reference <= int.MaxValue;
                    value = reference;
                    break;
                case "Date":
                    valid = DateOnly.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);
                    break;
                case "Decimal":
                    valid = PricePattern().IsMatch(text) && decimal.TryParse(text, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out _);
                    value = valid ? decimal.Parse(text, CultureInfo.InvariantCulture) : null;
                    break;
                case "Ubigeo":
                    valid=text.Length==6&&text.All(char.IsAsciiDigit);
                    break;
                case "Select":
                    valid = field.Catalog switch
                    {
                        "Currency" => text is "PEN" or "USD",
                        "TaxTreatment" => text is "CON_IGV" or "SIN_IGV",
                        "EntityType" => text is "JURIDICA" or "NATURAL",
                        _ => false
                    };
                    break;
            }
            if (!valid) issues.Add(new(field.Name, $"Revisa {field.Label.ToLowerInvariant()}."));
            values[field.Name] = value;
        }
        if(entity.Key is "clientes" or "tiendas")
        {
            var address=fields.GetValueOrDefault(entity.Key=="clientes"?"DireccionFiscal":"Direccion");
            if(!string.IsNullOrWhiteSpace(address)&&string.IsNullOrWhiteSpace(fields.GetValueOrDefault("CodigoUbigeo")))issues.Add(new("CodigoUbigeo","Selecciona departamento, provincia y distrito de la dirección."));
        }
        string? Text(string key) => values.GetValueOrDefault(key)?.ToString();
        if (entity.Key == "clientes")
        {
            var names = Text("TipoEntidad") == "NATURAL" ? new[] { "Nombres", "Apellidos" } : new[] { "RazonSocial" };
            foreach (var name in names)
                if (string.IsNullOrWhiteSpace(Text(name))) issues.Add(new(name, "Completa el nombre según el tipo de cliente."));
            var links=(documentLinks??[]).ToArray();
            if(links.Length>100)issues.Add(new("DocumentLinks","Puedes registrar hasta cien enlaces por cliente."));
            var normalized=new List<object>();
            for(var index=0;index<links.Length;index++)
            {
                var link=links[index];var description=link.Description?.Trim();var url=link.Url?.Trim();
                if(link.Id is null && string.IsNullOrEmpty(description) && string.IsNullOrEmpty(url))continue;
                if(link.Id is <=0)issues.Add(new($"DocumentLinks[{index}].Id","Vuelve a abrir la ficha para editar sus enlaces."));
                if(string.IsNullOrEmpty(description)||description.Length>200)issues.Add(new($"DocumentLinks[{index}].Description","Completa la descripción del enlace (máximo 200 caracteres)."));
                if(string.IsNullOrEmpty(url)||url.Length>2048||url.Any(char.IsControl)||!Uri.TryCreate(url,UriKind.Absolute,out var uri)||uri.Scheme is not ("http" or "https")||!string.IsNullOrEmpty(uri.UserInfo))
                    issues.Add(new($"DocumentLinks[{index}].Url","Ingresa un enlace HTTP o HTTPS válido, sin credenciales."));
                normalized.Add(new{link.Id,Descripcion=description,Url=url});
            }
            if(links.Where(link=>link.Id.HasValue).GroupBy(link=>link.Id).Any(group=>group.Count()>1))issues.Add(new("DocumentLinks","Revisa los enlaces repetidos."));
            values["Enlaces"]=normalized;
        }
        if (Text("Correo") is { Length: > 0 } email && !System.Net.Mail.MailAddress.TryCreate(email, out _))
            issues.Add(new("Correo", "Ingresa un correo válido."));
        if (DateOnly.TryParseExact(Text("VigenciaDesde"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start) &&
            DateOnly.TryParseExact(Text("VigenciaHasta"), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end) && end < start)
            issues.Add(new("VigenciaHasta", "La fecha final no puede ser anterior a la inicial."));
        if (entity.HasZones)
        {
            var selected = zones.Distinct().ToArray();
            if (selected.Any(zone => zone <= 0)) issues.Add(new("ZonaIds", "Revisa las zonas seleccionadas."));
            values["ZonaIds"] = selected;
        }
        return new(JsonSerializer.Serialize(values), version, issues);
    }
}
