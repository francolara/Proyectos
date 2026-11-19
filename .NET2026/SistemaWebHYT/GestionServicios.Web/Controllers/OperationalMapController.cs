using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

[Authorize]
[Route("MapaOperativo")]
public sealed class OperationalMapController(IWebHostEnvironment environment) : Controller
{
    [HttpGet]
    public IActionResult Index() => View();

    [HttpGet("Diagrama/{area}")]
    public IActionResult Diagram(string area)
    {
        if (area == "administracion")
        {
            return RedirectToAction(nameof(Diagram), new { area = "operacion", present = Request.Query["present"].ToString(), theme = Request.Query["theme"].ToString() });
        }

        if (area is not ("operacion" or "planillas"))
        {
            return NotFound();
        }

        var file = environment.ContentRootFileProvider.GetFileInfo($"OperationalMaps/{area}.html");
        return file.Exists ? File(file.CreateReadStream(), "text/html; charset=utf-8") : NotFound();
    }
}
