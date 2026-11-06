using System.Net.Http.Json;
using System.Net;
using System.Text.RegularExpressions;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Testing;

static class CaptchaChecks
{
    static string Code(string svg) => string.Concat(Regex.Matches(svg, @"<text[^>]*>([^<])</text>").Select(match => match.Groups[1].Value));
    static string Input(string html, string name)
    {
        var element=Regex.Match(html,"<input[^>]*name=\""+name+"\"[^>]*>").Value;
        return WebUtility.HtmlDecode(Regex.Match(element,"value=\"([^\"]*)\"").Groups[1].Value);
    }
    public static async Task Run(Action<bool,string> check)
    {
        var clock=new TestClock();using var service=new ManualCaptchaService(clock);
        var browser=Guid.NewGuid().ToString("N");var context=new DefaultHttpContext();context.Request.Headers.Cookie="__Host-Hyt.Captcha="+browser;
        var id=service.Create(context);var code=Code(service.Image(context,id)!);
        check(code.Length==6&&service.Validate(context,id,code.ToLowerInvariant()),"Captcha manual admite respuesta correcta sin distinguir mayúsculas");
        check(!service.Validate(context,id,code),"Captcha manual consumido no permite repetición");
        id=service.Create(context);code=Code(service.Image(context,id)!);
        check(!service.Validate(context,id,"000000")&&!service.Validate(context,id,code),"Respuesta errónea consume el intento de captcha");
        id=service.Create(context);code=Code(service.Image(context,id)!);
        var stranger=new DefaultHttpContext();stranger.Request.Headers.Cookie="__Host-Hyt.Captcha="+Guid.NewGuid().ToString("N");
        check(service.Image(stranger,id) is null&&!service.Validate(stranger,id,code)&&service.Validate(context,id,code),"Otro navegador no obtiene ni consume un captcha ajeno");
        id=service.Create(context);code=Code(service.Image(context,id)!);clock.Advance(TimeSpan.FromMinutes(6));
        check(service.Image(context,id) is null&&!service.Validate(context,id,code),"Captcha manual rechaza imágenes y respuestas vencidas");clock.Reset();
        id=service.Create(context);code=Code(service.Image(context,id)!);
        var attempts=await Task.WhenAll(Enumerable.Range(0,2).Select(_=>Task.Run(()=>service.Validate(context,id,code))));
        check(attempts.Count(valid=>valid)==1,"Dos intentos concurrentes solo consumen un captcha una vez");
        var fresh=new DefaultHttpContext();service.Create(fresh);var cookie=fresh.Response.Headers.SetCookie.ToString();
        check(cookie.Contains("secure",StringComparison.OrdinalIgnoreCase)&&cookie.Contains("httponly",StringComparison.OrdinalIgnoreCase)&&cookie.Contains("path=/")&&!cookie.Contains(code),"Captcha usa cookie de navegador segura sin guardar el código en ella");
        check(!service.Validate(context,null,"AAAAAA"),"Marcar modo manual sin desafío no omite la validación");

        var repository=new TestAccessRepository();using var factory=new HytFactory(repository,true,"");
        using var client=factory.CreateClient(new WebApplicationFactoryClientOptions{BaseAddress=new Uri("https://localhost"),AllowAutoRedirect=false});
        var html=await client.GetStringAsync("/Acceso");
        check(html.Contains("Verificación manual")&&Input(html,"UseManualCaptcha")=="True"&&!html.Contains("disabled=\"disabled\">Ingresar"),"Sin claves públicas se habilita captcha manual e ingreso");
        check(html.Contains("hyt-loading.js")&&html.Contains("hyt-loading.css"),"Login incluye indicador compartido de espera");
        var csrf=Input(html,"__RequestVerificationToken");id=Guid.Parse(Input(html,"CaptchaId"));
        var image=await client.GetAsync("/Acceso/Captcha/"+id);
        check(image.StatusCode==HttpStatusCode.OK&&image.Content.Headers.ContentType?.MediaType=="image/svg+xml"&&image.Headers.CacheControl?.NoStore==true,"Imagen de captcha se sirve sin caché y vinculada al navegador");
        code=Code(await image.Content.ReadAsStringAsync());
        check(!html.Contains(code),"El HTML no expone la solución del captcha");
        async Task<HttpResponseMessage> Login(Guid? challenge,string answer,string? password=null)=>await client.PostAsync("/Acceso",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=csrf,["UserName"]="operador",["Password"]=password??repository.Password,["UseManualCaptcha"]="true",["CaptchaId"]=challenge?.ToString()??"",["CaptchaAnswer"]=answer}));
        var wrong=await Login(id,"000000");html=await wrong.Content.ReadAsStringAsync();
        check(wrong.StatusCode==HttpStatusCode.OK&&repository.NameLookups==0&&WebUtility.HtmlDecode(html).Contains("incorrecto o venció"),"Captcha erróneo no consulta credenciales y ofrece un código nuevo");
        check(Input(html,"CaptchaId")!=id.ToString()&&string.IsNullOrEmpty(Input(html,"Password")),"Rechazo renueva desafío y no devuelve contraseña en HTML");
        var forged=await Login(null,"AAAAAA");
        check(forged.StatusCode==HttpStatusCode.OK&&repository.NameLookups==0,"Modo manual falsificado sin desafío no autentica");
        html=await forged.Content.ReadAsStringAsync();id=Guid.Parse(Input(html,"CaptchaId"));csrf=Input(html,"__RequestVerificationToken");
        code=Code(await client.GetStringAsync("/Acceso/Captcha/"+id));
        var badPassword=await Login(id,code,repository.Password+"incorrecta");html=await badPassword.Content.ReadAsStringAsync();
        check(badPassword.StatusCode==HttpStatusCode.OK&&repository.NameLookups==1&&repository.Failures==1&&!badPassword.Headers.TryGetValues("Set-Cookie",out var badCookies),"Resolver captcha no evita validar contraseña ni registra sesión si es incorrecta");
        id=Guid.Parse(Input(html,"CaptchaId"));csrf=Input(html,"__RequestVerificationToken");code=Code(await client.GetStringAsync("/Acceso/Captcha/"+id));
        var success=await Login(id,code);
        check(success.StatusCode==HttpStatusCode.Redirect&&repository.NameLookups==2,"Captcha manual válido permite verificar contraseña y emitir sesión");

        var rejectedRepository=new TestAccessRepository();using var rejectedFactory=new HytFactory(rejectedRepository,true);
        using var rejectedClient=rejectedFactory.CreateClient(new WebApplicationFactoryClientOptions{BaseAddress=new Uri("https://localhost"),AllowAutoRedirect=false});
        html=await rejectedClient.GetStringAsync("/Acceso");csrf=Input(html,"__RequestVerificationToken");
        var unavailable=await rejectedClient.PostAsync("/Acceso",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=csrf,["UserName"]="operador",["Password"]=rejectedRepository.Password}));
        html=await unavailable.Content.ReadAsStringAsync();
        check(Input(html,"UseManualCaptcha")=="True"&&rejectedRepository.NameLookups==0,"Fallo de Turnstile activa respaldo sin aceptar todavía credenciales");
        var another=await rejectedClient.GetFromJsonAsync<System.Text.Json.JsonElement>("/Acceso/Captcha/Nuevo");
        check(another.GetProperty("id").GetGuid()!=Guid.Parse(Input(html,"CaptchaId")),"Otro código entrega un identificador nuevo sin exponer respuesta");
    }
}
