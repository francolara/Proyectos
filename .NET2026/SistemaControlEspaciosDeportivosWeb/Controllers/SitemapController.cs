using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using System.Xml.Linq;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[ApiExplorerSettings(IgnoreApi = true)]
public class SitemapController(ISportCenterStoredProcedureService spService) : Controller
{
    private const string OfficialBaseUrl = "https://www.lazonadeportiva.com";

    [HttpGet("/sitemap.xml")]
    [ResponseCache(Duration = 3600, Location = ResponseCacheLocation.Any)]
    public async Task<IActionResult> Index()
    {
        var hoy = DateTime.UtcNow.ToString("yyyy-MM-dd");
        var ns = XNamespace.Get("http://www.sitemaps.org/schemas/sitemap/0.9");
        var sedes = await spService.HomeListarSedesAsync();
        var complejos = ComplejosController.ConstruirComplejos(sedes);
        var boletines = await spService.BoletinesDeportivosListarPublicoAsync();

        var urls = new List<XElement>
        {
            CrearUrl(ns, OfficialBaseUrl + "/", hoy, "daily", "1.0"),
            CrearUrl(ns, OfficialBaseUrl + "/Home/Faq", hoy, "monthly", "0.7"),
            CrearUrl(ns, OfficialBaseUrl + "/complejos", hoy, "daily", "0.9"),
            CrearUrl(ns, OfficialBaseUrl + "/Boletines", hoy, "daily", "0.8")
        };

        urls.AddRange(complejos.Select(complejo =>
            CrearUrl(
                ns,
                $"{OfficialBaseUrl}/complejos/{complejo.Slug}-{complejo.NegocioId}",
                hoy,
                "weekly",
                "0.8")));

        urls.AddRange(boletines.Select(boletin =>
            CrearUrl(
                ns,
                $"{OfficialBaseUrl}/Boletines/Detalle/{boletin.IdBoletin}",
                boletin.FechaCreacion == DateTime.MinValue ? hoy : boletin.FechaCreacion.ToUniversalTime().ToString("yyyy-MM-dd"),
                "weekly",
                "0.7")));

        var document = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "urlset", urls));

        return Content(document.ToString(SaveOptions.DisableFormatting), "application/xml; charset=utf-8");
    }

    private static XElement CrearUrl(XNamespace ns, string loc, string lastmod, string changefreq, string priority)
    {
        return new XElement(ns + "url",
            new XElement(ns + "loc", loc),
            new XElement(ns + "lastmod", lastmod),
            new XElement(ns + "changefreq", changefreq),
            new XElement(ns + "priority", priority));
    }
}
