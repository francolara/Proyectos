using Microsoft.AspNetCore.Razor.TagHelpers;

namespace GestionServicios.Web.TagHelpers;

// Firma: FRANCO LARA - 11/11/2026 | Identifica las acciones de consulta y edición para usar el estilo compacto común.
[HtmlTargetElement("a", Attributes = "class")]
[HtmlTargetElement("button", Attributes = "class")]
public sealed class CompactRecordActionTagHelper : TagHelper
{
    public override async Task ProcessAsync(TagHelperContext context, TagHelperOutput output)
    {
        var classes = output.Attributes["class"]?.Value?.ToString()?.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (classes is null || !classes.Any(name => name is "hyt-button-primary" or "hyt-button-secondary"
            or "hyt-action-button" or "hyt-record-action"))
            return;

        var content = output.Content.IsModified ? output.Content : await output.GetChildContentAsync();
        var label = content.GetContent().Trim();
        if (HasAction(label, "Ver"))
            output.Attributes.SetAttribute("data-hyt-record-action", "view");
        else if (HasAction(label, "Editar"))
            output.Attributes.SetAttribute("data-hyt-record-action", "edit");
    }

    private static bool HasAction(string label, string action) =>
        label.Equals(action, StringComparison.OrdinalIgnoreCase) ||
        (label.Length > action.Length && char.IsWhiteSpace(label[action.Length]) &&
            label.StartsWith(action, StringComparison.OrdinalIgnoreCase));
}
