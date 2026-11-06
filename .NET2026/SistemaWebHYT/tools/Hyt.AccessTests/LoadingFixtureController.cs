using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using GestionServicios.Web.Security;
using GestionServicios.Application.Security;

// Solo en el ensamblado de pruebas; permite observar respuestas demoradas y errores sin SQL.
[AllowAnonymous]
public sealed class LoadingFixtureController : Controller
{
    private static int posts;
    private static string? intent;
    [HttpGet("/__tests/loading-count")]
    public IActionResult Count([FromServices] IAccessRepository repository)=>Json(new{posts,intent,lookups=((TestAccessRepository)repository).NameLookups});
    [HttpGet("/__tests/challenge-mode")]
    public IActionResult ChallengeMode([FromServices]IOptions<CloudflareTurnstileSettings> options,string mode="manual")
    {
        options.Value.Enabled=true;options.Value.SiteKey=mode=="widget"?"sitekey-prueba":"";
        return Json(new{ok=true});
    }
    [HttpGet("/__tests/loading-delay")]
    public async Task<IActionResult> Delay(int ms=700,string kind="json",CancellationToken ct=default)
    {
        await Task.Delay(Math.Clamp(ms,0,5000),ct);
        return kind switch{"error"=>StatusCode(503),"invalid"=>Content("no es JSON","application/json"),"file"=>File("%PDF-1.7\nPrueba"u8.ToArray(),"application/pdf","guia-prueba.pdf"),_=>Json(new{ok=true})};
    }
    [HttpGet("/__tests/loading-page")]
    public async Task<IActionResult> Page(int ms=700,CancellationToken ct=default){await Task.Delay(Math.Clamp(ms,0,5000),ct);return View("~/Views/Home/Index.cshtml");}
    [HttpGet("/__tests/loading-body")]
    public async Task Body(CancellationToken ct)
    {
        Response.ContentType="application/json";await Response.StartAsync(ct);await Response.Body.FlushAsync(ct);
        await Task.Delay(900,ct);await Response.WriteAsync("{\"ok\":true}",ct);
    }
    [HttpPost("/__tests/loading-submit")]
    public async Task<IActionResult> Submit([FromForm]string? intent,CancellationToken ct)
    {
        Interlocked.Increment(ref posts);LoadingFixtureController.intent=intent;
        await Task.Delay(700,ct);return Redirect("/__tests/loading-page?ms=0");
    }
}
