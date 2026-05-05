using Microsoft.AspNetCore.Mvc;
using System.Xml.Linq;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[ApiExplorerSettings(IgnoreApi = true)]
public class SitemapController : Controller
{
    [HttpGet("/sitemap.xml")]
    [ResponseCache(Duration = 3600, Location = ResponseCacheLocation.Any)]
    public async Task<IActionResult> Index()
    {
        var hoy = DateTime.UtcNow.ToString("yyyy-MM-dd");
        var ns = XNamespace.Get("http://www.sitemaps.org/schemas/sitemap/0.9");

        var urls = new List<XElement>
        {
            CrearUrl(ns, Url.Action("Index", "Home", null, Request.Scheme)!, hoy, "daily", "1.0"),
            CrearUrl(ns, Url.Action("Index", "Complejos", null, Request.Scheme)!, hoy, "daily", "0.9")
        };

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
