using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Text;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[Route("complejos")]
public class ComplejosController(ISportCenterStoredProcedureService spService) : Controller
{
    [HttpGet("")]
    public async Task<IActionResult> Index()
    {
        var sedes = await spService.HomeListarSedesAsync();
        var complejos = ConstruirComplejos(sedes);
        var vm = new ComplejosIndexViewModel { Complejos = complejos };
        return View(vm);
    }

    [HttpGet("{slug}-{negocioId:int}")]
    public async Task<IActionResult> Detalle(string slug, int negocioId)
    {
        var sedes = await spService.HomeListarSedesAsync();
        var complejos = ConstruirComplejos(sedes);
        var complejo = complejos.FirstOrDefault(x => x.NegocioId == negocioId);
        if (complejo is null)
            return NotFound();

        if (!string.Equals(slug, complejo.Slug, StringComparison.OrdinalIgnoreCase))
            return RedirectToActionPermanent(nameof(Detalle), new { slug = complejo.Slug, negocioId });

        var sedesComplejo = sedes
            .Where(x => x.NegocioId == negocioId)
            .OrderBy(x => x.Nombre)
            .ToList();

        var vm = new ComplejoDetalleViewModel
        {
            NegocioId = negocioId,
            Nombre = complejo.Nombre,
            Slug = complejo.Slug,
            Sedes = sedesComplejo
        };

        return View(vm);
    }

    internal static List<ComplejoListadoItemViewModel> ConstruirComplejos(List<SedePublicaViewModel> sedes)
    {
        return sedes
            .Where(x => x.NegocioId.HasValue)
            .GroupBy(x => x.NegocioId!.Value)
            .Select(g =>
            {
                var nombre = g.Select(x => x.NegocioNombre)
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))
                    ?.Trim() ?? $"Complejo {g.Key}";
                return new ComplejoListadoItemViewModel
                {
                    NegocioId = g.Key,
                    Nombre = nombre,
                    Slug = ToSlug(nombre),
                    TotalSedes = g.Count(),
                    DireccionReferencia = g.Select(x => x.Direccion).FirstOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim()
                };
            })
            .OrderBy(x => x.Nombre)
            .ToList();
    }

    internal static string ToSlug(string? value)
    {
        var input = (value ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(input))
            return "complejo";

        var sb = new StringBuilder(input.Length);
        foreach (var ch in input.Normalize(NormalizationForm.FormD))
        {
            if (char.GetUnicodeCategory(ch) == System.Globalization.UnicodeCategory.NonSpacingMark)
                continue;
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(ch);
                continue;
            }
            if (ch is ' ' or '-' or '_' or '.')
                sb.Append('-');
        }

        var slug = sb.ToString().Trim('-');
        while (slug.Contains("--", StringComparison.Ordinal))
            slug = slug.Replace("--", "-", StringComparison.Ordinal);
        return string.IsNullOrWhiteSpace(slug) ? "complejo" : slug;
    }
}
