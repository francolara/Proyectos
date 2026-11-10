using System.Net;
using System.Text.RegularExpressions;
using GestionServicios.Application.Security;
using GestionServicios.Web;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.AspNetCore.Mvc.Controllers;
using GestionServicios.Application.Maintenance;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Payroll;

string[] payrollPermissions = ["PROVEEDORES.CONSULTAR", "PROVEEDORES.EDITAR", "PROVEEDORES.CUENTA",
    "COSTOS_PROVEEDOR.CONSULTAR", "COSTOS_PROVEEDOR.EDITAR", "PAGOS_PROVEEDOR.CONSULTAR", "PAGOS_PROVEEDOR.EDITAR",
    "TRABAJADORES.CONSULTAR", "TRABAJADORES.EDITAR", "TRABAJADORES.CUENTA", "VACACIONES.CONSULTAR", "VACACIONES.EDITAR",
    "ASISTENCIA.CONSULTAR", "ASISTENCIA.EDITAR", "PAGOS_PERSONAL.CONSULTAR", "PAGOS_PERSONAL.EDITAR",
    "PAGOS_PERSONAL.CESES", "PAGOS_PERSONAL.CERRAR", "REPORTES_PLANILLAS.CONSULTAR",
    "OBLIGACIONES.CONSULTAR", "OBLIGACIONES.EDITAR"];

if (args.Contains("--browser-host"))
{
    var browserRepository = new TestAccessRepository { LookupDelay=args.Contains("--browser-captcha")?900:0 };
    browserRepository.Account = browserRepository.Account with { Permissions = MaintenanceDefinitions.All.SelectMany(entity => new[] { entity.Permission + ".CONSULTAR", entity.Permission + ".EDITAR" }).Concat(["CONFIGURACION.CONSULTAR","CONFIGURACION.EDITAR","PRESUPUESTOS.CONSULTAR","PRESUPUESTOS.EDITAR","REPORTES.CONSULTAR","OC_CLIENTE.CONSULTAR","GUIAS.CONSULTAR","CONFIRMACIONES.CONSULTAR","EJECUCION.CONSULTAR","OC_CLIENTE.EDITAR","GUIAS.EDITAR","CONFIRMACIONES.EDITAR","EJECUCION.EDITAR","FACTURADO.CONSULTAR","FACTURADO.EDITAR"]).Distinct().ToArray() };
    if (args.Contains("--browser-administration"))
        browserRepository.Account = browserRepository.Account with
        {
            Permissions = browserRepository.Account.Permissions.Concat(["USUARIOS.CONSULTAR", "USUARIOS.CREAR",
                "USUARIOS.MODIFICAR", "USUARIOS.RESTABLECER", "PERFILES.CONSULTAR", "PERFILES.EDITAR", "AUDITORIA.CONSULTAR"]).ToArray(),
            MustChangePassword = args.Contains("--browser-temporary-password")
        };
    if (args.Contains("--browser-payroll"))
        browserRepository.Account = browserRepository.Account with
        {
            Permissions = browserRepository.Account.Permissions.Concat(payrollPermissions).Distinct().ToArray()
        };
    using var browserFactory = new HytFactory(browserRepository,args.Contains("--browser-captcha"),args.Contains("--browser-captcha")?"":"sitekey-prueba");
    browserFactory.UseKestrel(5198);
    using var browserClient = browserFactory.CreateClient();
    if(args.Contains("--browser-budget-documents")){browserFactory.Budgets.RegisteredDocuments=true;browserFactory.Documents.RichList=true;}
    if(args.Contains("--browser-oc-details"))
    {
        browserFactory.Documents.RichList=true;
        await browserFactory.InvoiceMarks.SaveAsync("browser-test",new(1,new byte[8],Guid.NewGuid(),new DateOnly(2026,11,3),"01","F001","000012A",null,null,false,null),CancellationToken.None);
    }
    var fixturePath = Path.Combine(Path.GetTempPath(), "hyt-maintenance-browser-fixture.json");
    await File.WriteAllTextAsync(fixturePath, System.Text.Json.JsonSerializer.Serialize(new { url = "http://localhost:5198", user = "operador", password = browserRepository.Password }));
    Console.WriteLine("Host de prueba MVC local preparado, sin SQL. Pulsa Enter para cerrar.");
    Console.ReadLine();
    File.Delete(fixturePath);
    return;
}

void ValidateMaintenance(bool condition, string label)
{
    if (!condition) throw new InvalidOperationException(label);
    Console.WriteLine("OK: " + label);
}
var netLine = BudgetAmounts.Calculate(2, 100, "PEN", "SIN_IGV", "PEN", "SIN_IGV", 0.18m);
ValidateMaintenance(netLine.UnitPrice == 100 && netLine.Net == 200 && netLine.Tax == 36 && netLine.Total == 236, "Presupuesto SIN IGV calcula base, impuesto y total por partida");
var grossLine = BudgetAmounts.Calculate(2, 118, "PEN", "CON_IGV", "PEN", "CON_IGV", 0.18m);
ValidateMaintenance(grossLine.Net == 200 && grossLine.Tax == 36 && grossLine.Total == 236, "Precio CON IGV no vuelve a sumar impuesto");
var convertedLine = BudgetAmounts.Calculate(1, 118, "PEN", "CON_IGV", "PEN", "SIN_IGV", 0.18m);
ValidateMaintenance(convertedLine.UnitPrice == 100 && convertedLine.Total == 118, "Conversión entre tratamientos IGV mantiene importes equivalentes");
var missingPrice = BudgetAmounts.Calculate(1, null, "USD", "SIN_IGV", "USD", "SIN_IGV", 0.18m);
ValidateMaintenance(missingPrice.Total is null && !BudgetAmounts.Sum([missingPrice]).Complete, "USD ausente deja el presupuesto incompleto sin convertirlo en cero");
ValidateMaintenance(BudgetAmounts.Calculate(1, 0.03m, "PEN", "SIN_IGV", "PEN", "SIN_IGV", 0.18m).Tax == 0.01m && BudgetAmounts.Sum([netLine, convertedLine]).Total == 354, "Redondeo comercial a dos decimales y suma de partidas");
var rejectedDecimals = false;
try { BudgetAmounts.Calculate(1.001m, 1, "PEN", "SIN_IGV", "PEN", "SIN_IGV", 0.18m); } catch (ArgumentException) { rejectedDecimals = true; }
ValidateMaintenance(rejectedDecimals, "Cantidad con más de dos decimales se rechaza");
var rejectedCurrency = false;
try { BudgetAmounts.Calculate(1, 10, "PEN", "SIN_IGV", "USD", "SIN_IGV", 0.18m); } catch (ArgumentException) { rejectedCurrency = true; }
ValidateMaintenance(rejectedCurrency, "Cambio de moneda exige revalorizar y no aplica tipo de cambio automático");
var materialDefinition = MaintenanceDefinitions.Find("materiales")!;
var materialFields = new Dictionary<string, string?>
{
    ["CodigoInterno"] = "0001-A", ["Descripcion"] = "Servicio", ["UnidadMedidaId"] = "1", ["Activo"] = "true"
};
var materialInput = MaintenanceInput.Validate(materialDefinition, materialFields, null, null, [1, 2, 1]);
ValidateMaintenance(materialInput.Issues.Count == 0 && !materialInput.Payload.Contains("0001-A") && materialInput.Payload.Contains("[1,2]"), "Material ignora código interno enviado y conserva zonas sin duplicados");
ValidateMaintenance(MaintenanceInput.Validate(materialDefinition, materialFields, 1, "inválida", []).Issues.Count > 0, "Edición rechaza versión de concurrencia inválida");
var priceDefinition = MaintenanceDefinitions.Find("partidas-tarifario")!;
var priceFields = new Dictionary<string, string?>
{
    ["VersionId"] = "1", ["MaterialId"] = "2", ["CodigoMoneda"] = "PEN", ["CodigoTipoIGV"] = "SIN_IGV", ["PrecioPEN"] = "12.12", ["Activo"] = "true"
};
var priceInput = MaintenanceInput.Validate(priceDefinition, priceFields, null, null, []);
ValidateMaintenance(priceInput.Issues.Count == 0 && priceInput.Payload.Contains("\"PrecioUSD\":null"), "Precio USD ausente se conserva como nulo");
priceFields["PrecioPEN"] = "12.123";
ValidateMaintenance(MaintenanceInput.Validate(priceDefinition, priceFields, null, null, []).Issues.Any(issue => issue.Field == "PrecioPEN"), "Precio con más de dos decimales requiere corrección");
priceFields["PrecioPEN"] = "-1";
ValidateMaintenance(MaintenanceInput.Validate(priceDefinition, priceFields, null, null, []).Issues.Any(issue => issue.Field == "PrecioPEN"), "Precio negativo rechazado");
var clientFields = new Dictionary<string, string?> { ["Codigo"] = "C1", ["TipoEntidad"] = "NATURAL", ["Activo"] = "true", ["UrlCargaDocumentos"] = "javascript:alert(1)" };
var clientIssues = MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!, clientFields, null, null, [], [new(){Description="Portal",Url="javascript:alert(1)"}]).Issues;
ValidateMaintenance(clientIssues.Any(issue => issue.Field == "Nombres") && clientIssues.Any(issue => issue.Field == "Apellidos") && clientIssues.Any(issue => issue.Field == "DocumentLinks[0].Url"), "Cliente natural exige nombres y rechaza enlace ejecutable");

clientFields["UrlCargaDocumentos"]="";clientFields["Nombres"]="Ana";clientFields["Apellidos"]="Pérez";clientFields["TipoIdentificacion"]="CE";clientFields["NumeroIdentificacion"]="CE123456";clientFields["RazonSocial"]="No guardar";
var naturalClient=MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!,clientFields,null,null,[]);
ValidateMaintenance(naturalClient.Issues.Count==0&&naturalClient.Payload.Contains("\"RazonSocial\":null"),"Natural admite CE y descarta razón social");
clientFields["TipoIdentificacion"]="RUC";
ValidateMaintenance(!MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!,clientFields,null,null,[]).Issues.Any(issue=>issue.Field=="TipoIdentificacion"),"Tipo oculto no valida documentos de persona natural");
clientFields["TipoEntidad"]="JURIDICA";
var legalClient=MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!,clientFields,null,null,[]);
ValidateMaintenance(legalClient.Issues.Count==0&&legalClient.Payload.Contains("\"Nombres\":null")&&legalClient.Payload.Contains("\"Apellidos\":null"),"Jurídica admite RUC y descarta nombres y apellidos");
clientFields["TipoIdentificacion"]="DNI";
ValidateMaintenance(!MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!,clientFields,null,null,[]).Issues.Any(issue=>issue.Field=="TipoIdentificacion"),"Tipo oculto no valida documentos de persona jurídica");
clientFields["TipoIdentificacion"]="RUC";clientFields["NumeroIdentificacion"]="";
ValidateMaintenance(MaintenanceInput.Validate(MaintenanceDefinitions.Find("clientes")!,clientFields,null,null,[]).Issues.Any(issue=>issue.Field=="NumeroIdentificacion"),"Cliente exige número de documento");

var repository = new TestAccessRepository();
using var factory = new HytFactory(repository);
using var client = factory.CreateClient(new WebApplicationFactoryClientOptions { BaseAddress = new Uri("https://localhost"), AllowAutoRedirect = false });
void Check(bool condition, string label) { if (!condition) throw new InvalidOperationException(label); Console.WriteLine("OK: " + label); }
async Task<string> Token(string path)
{
    var html = await client.GetStringAsync(path);
    var input = Regex.Match(html, "<input[^>]*name=\"__RequestVerificationToken\"[^>]*>").Value;
    return WebUtility.HtmlDecode(Regex.Match(input, "value=\"([^\"]+)\"").Groups[1].Value);
}
async Task<HttpResponseMessage> Login(string name, string password, string? returnUrl = null)
{
    var token = await Token("/Acceso");
    return await client.PostAsync("/Acceso", new FormUrlEncodedContent(new Dictionary<string, string>
    { ["UserName"] = name, ["Password"] = password, ["ReturnUrl"] = returnUrl ?? "/", ["__RequestVerificationToken"] = token }));
}

Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Inicio privado redirige al acceso");
var loginHtml = await client.GetStringAsync("/Acceso");
Check(loginHtml.Contains("HYT") && loginHtml.Contains("autocomplete=\"current-password\""), "Acceso HYT y campos identificados");
Check((await client.PostAsync("/Acceso", new FormUrlEncodedContent(new Dictionary<string, string> { ["UserName"] = "operador", ["Password"] = repository.Password }))).StatusCode == HttpStatusCode.BadRequest, "POST sin antifalsificación rechazado");
var unknown = await Login("no-existe", "incorrecta");
Check(WebUtility.HtmlDecode(await unknown.Content.ReadAsStringAsync()).Contains("No se pudo iniciar sesión"), "Usuario desconocido recibe error genérico");
var bad = await Login("operador", "incorrecta");
Check(WebUtility.HtmlDecode(await bad.Content.ReadAsStringAsync()).Contains("No se pudo iniciar sesión") && repository.Failures == 1, "Contraseña incorrecta registra intento fallido");
repository.Account = repository.Account with { Active = false };
Check(WebUtility.HtmlDecode(await (await Login("operador", repository.Password)).Content.ReadAsStringAsync()).Contains("No se pudo iniciar sesión"), "Cuenta desactivada no puede entrar");
repository.Account = repository.Account with { Active = true, LockoutEnd = DateTimeOffset.UtcNow.AddMinutes(15) };
Check(WebUtility.HtmlDecode(await (await Login("operador", repository.Password)).Content.ReadAsStringAsync()).Contains("No se pudo iniciar sesión"), "Cuenta bloqueada no puede entrar");
repository.Account = repository.Account with { LockoutEnd = null };
var success = await Login("operador", repository.Password, "https://externo.example/");
Check(success.StatusCode == HttpStatusCode.Redirect && success.Headers.Location?.OriginalString == "/", "Acceso válido y rechazo de retorno externo");
Check(success.Headers.GetValues("Set-Cookie").Any(value => value.Contains("__Host-Hyt.Session") && value.Contains("secure", StringComparison.OrdinalIgnoreCase) && value.Contains("httponly", StringComparison.OrdinalIgnoreCase)), "Cookie segura, privada y con prefijo Host");
var home = await client.GetAsync("/");
var homeHtml = await home.Content.ReadAsStringAsync();
Check(home.StatusCode == HttpStatusCode.OK && homeHtml.Contains("/modulo/clientes"), "Sesión válida abre inicio y menú permitido");
Check(!homeHtml.Contains("Usuarios y permisos"), "Menú omite opción sin permiso");
var deniedResponse = await client.GetAsync("/modulo/usuarios");
var denied = await deniedResponse.Content.ReadAsStringAsync();
Check(deniedResponse.StatusCode == HttpStatusCode.Forbidden && denied.Contains("Acceso restringido") && !denied.Contains("Módulo en preparación"), "Ruta de módulo sin permiso no muestra contenido protegido");
repository.Account = repository.Account with { Active = false };
Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Desactivación invalida sesión existente en siguiente solicitud");
repository.Account = repository.Account with { Active = true };
await Login("operador", repository.Password);
repository.Account = repository.Account with { SecurityStamp = Guid.NewGuid().ToString("N") };
Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Cambio de sello invalida sesión existente");
await Login("operador", repository.Password);
Check((await client.GetAsync("/Salir")).StatusCode == HttpStatusCode.OK && (await client.GetAsync("/")).StatusCode == HttpStatusCode.OK, "GET de salida solicita confirmación y no cierra sesión");
var logoutToken = await Token("/Salir");
var logout = await client.PostAsync("/Salir", new FormUrlEncodedContent(new Dictionary<string, string> { ["__RequestVerificationToken"] = logoutToken }));
Check(logout.StatusCode == HttpStatusCode.Redirect && repository.Logouts == 1, "Cierre POST registrado en auditoría");
Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Sesión cerrada ya no accede al inicio");
await Login("operador", repository.Password);
repository.Account = repository.Account with { Permissions = [] };
Check(!(await client.GetStringAsync("/")).Contains("Clientes y tiendas"), "Permiso retirado desaparece en la siguiente solicitud");
repository.Unavailable = true;
Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Fallo del repositorio rechaza una sesión existente");
repository.Unavailable = false;
await Login("operador", repository.Password);
factory.Clock.Advance(TimeSpan.FromHours(8) + TimeSpan.FromMinutes(1));
Check((await client.GetAsync("/")).StatusCode == HttpStatusCode.Redirect, "Límite absoluto de ocho horas invalida la sesión");
Check(factory.Services.GetRequiredService<IActionDescriptorCollectionProvider>().ActionDescriptors.Items.All(item => item is ControllerActionDescriptor), "Todas las acciones activas pertenecen a controladores MVC");
Check(!homeHtml.Contains("_framework/blazor") && !homeHtml.Contains("blazor.web.js"), "Las vistas no cargan el runtime Blazor");
var gridHtml = WebUtility.HtmlDecode(await client.GetStringAsync("/__tests/parts/grid"));
Check(gridHtml.Contains("20 por página") && gridHtml.Contains("page=2") && gridHtml.Contains("Search=buscar"), "Grilla MVC conserva filtros y página SQL de 20");
var rawGrid = await client.GetStringAsync("/__tests/parts/grid");
Check(!rawGrid.Contains("<prueba>"), "La grilla codifica datos de texto sin inyectar HTML");
var dateHtml = await client.GetStringAsync("/__tests/parts/date");
Check(dateHtml.Contains("type=\"date\"") && dateHtml.Contains("value=\"2026-11-03\"") && dateHtml.Contains("name=\"Fecha\""), "Fecha MVC conserva nombre y valor ISO para el controlador");
var filterHtml = await client.GetStringAsync("/__tests/parts/filters");
Check(filterHtml.Contains("name=\"Search\"") && filterHtml.Contains("name=\"From\"") && filterHtml.Contains("name=\"To\""), "Filtros MVC envían búsqueda y rango de fechas");
var traceHtml = WebUtility.HtmlDecode(await client.GetStringAsync("/__tests/parts/trace"));
Check(traceHtml.Contains("03/11/2026 13:30") && traceHtml.Contains("Sin actualizaciones"), "Auditoría MVC muestra fecha y hora de Perú");
var viewHtml = await client.GetStringAsync("/__tests/parts/view");
Check(viewHtml.Contains("<fieldset disabled") && !viewHtml.Contains("type=\"submit\""), "Ficha MVC de consulta no ofrece guardar ni editar campos");
var modalHtml = await client.GetStringAsync("/__tests/parts/modal");
Check(modalHtml.Contains("<dialog") && modalHtml.Contains("data-hyt-close"), "Ayuda MVC usa popup común accesible");
var securityRepository = new TestAccessRepository();
using var securityFactory = new HytFactory(securityRepository);
using var securityClient = securityFactory.CreateClient(new WebApplicationFactoryClientOptions
{
    BaseAddress = new Uri("https://localhost"), AllowAutoRedirect = false
});
async Task<string> SecurityToken(string path)
{
    var html = await securityClient.GetStringAsync(path);
    var input = Regex.Match(html, "<input[^>]*name=\"__RequestVerificationToken\"[^>]*>").Value;
    return WebUtility.HtmlDecode(Regex.Match(input, "value=\"([^\"]+)\"").Groups[1].Value);
}
securityRepository.Account = securityRepository.Account with
{
    MustChangePassword = true, SecurityStamp = Guid.NewGuid().ToString("N"), Permissions = ["CLIENTES.CONSULTAR", "CLIENTES.CREAR"]
};
var temporary = await securityClient.PostAsync("/Acceso", new FormUrlEncodedContent(new Dictionary<string, string>
{
    ["__RequestVerificationToken"] = await SecurityToken("/Acceso"), ["UserName"] = "operador",
    ["Password"] = securityRepository.Password, ["ReturnUrl"] = "/Maestros/clientes"
}));
Check(temporary.Headers.Location?.OriginalString == "/Cuenta/CambiarClave", "Contraseña temporal ignora retorno a módulo y exige cambio");
using (var navigation = new HttpRequestMessage(HttpMethod.Get, "/Maestros/clientes"))
{
    navigation.Headers.Accept.ParseAdd("text/html");
    Check((await securityClient.SendAsync(navigation)).Headers.Location?.OriginalString == "/Cuenta/CambiarClave",
        "Navegación directa queda restringida al cambio de contraseña");
}
Check((await securityClient.GetAsync("/Seguimiento/guias/Archivos/1?download=true")).StatusCode == HttpStatusCode.Forbidden,
    "Sesión temporal no puede obtener adjuntos por URL directa");
Check((await securityClient.PostAsync("/Maestros/clientes/Guardar", new FormUrlEncodedContent([]))).StatusCode == HttpStatusCode.Forbidden,
    "Sesión temporal no ejecuta operaciones de negocio");
Check((await securityClient.GetAsync("/hyt-theme.css")).StatusCode == HttpStatusCode.OK, "El formulario usa estilos públicos de HYT");
var passwordPage = await securityClient.GetStringAsync("/Cuenta/CambiarClave");
Check(!passwordPage.Contains("name=\"CurrentPassword\"") && passwordPage.Contains("name=\"NewPassword\""),
    "Primer ingreso solicita contraseña nueva y confirmación");
async Task<HttpResponseMessage> ChangePassword(string next, string? current = null, string? confirm = null) =>
    await securityClient.PostAsync("/Cuenta/CambiarClave", new FormUrlEncodedContent(new Dictionary<string, string>
    {
        ["__RequestVerificationToken"] = await SecurityToken("/Cuenta/CambiarClave"), ["NewPassword"] = next,
        ["ConfirmPassword"] = confirm ?? next, ["CurrentPassword"] = current ?? ""
    }));
Check((await ChangePassword(securityRepository.Password)).StatusCode == HttpStatusCode.OK && securityRepository.Account.MustChangePassword,
    "No se puede reutilizar la contraseña temporal");
var newPassword = Guid.NewGuid().ToString("N");
Check((await ChangePassword(newPassword, confirm: "diferente")).StatusCode == HttpStatusCode.OK && securityRepository.Account.MustChangePassword,
    "La confirmación debe coincidir");
Check((await ChangePassword(newPassword)).Headers.Location?.OriginalString == "/" && !securityRepository.Account.MustChangePassword,
    "Cambio obligatorio renueva la sesión y libera el acceso");
var createOnly = await securityClient.GetAsync("/Maestros/clientes/Nuevo");
var editDenied = await securityClient.GetAsync("/Maestros/clientes/1/Editar");
Check(createOnly.StatusCode == HttpStatusCode.OK &&
    (editDenied.StatusCode == HttpStatusCode.Forbidden || editDenied.Headers.Location?.OriginalString.Contains("/Denegado", StringComparison.Ordinal) == true),
    "Permiso para crear muestra el formulario nuevo e impide editar");
Check((await ChangePassword(Guid.NewGuid().ToString("N"), "incorrecta")).StatusCode == HttpStatusCode.OK,
    "Cambio voluntario requiere la contraseña actual correcta");
Check((await ChangePassword(Guid.NewGuid().ToString("N"), newPassword)).Headers.Location?.OriginalString == "/",
    "Usuario de consulta puede cambiar su propia contraseña");
securityRepository.Account = securityRepository.Account with { Permissions = ["USUARIOS.CONSULTAR", "USUARIOS.CREAR", "USUARIOS.MODIFICAR", "PERFILES.CONSULTAR", "PERFILES.EDITAR", "AUDITORIA.CONSULTAR"] };
Check((await securityClient.GetAsync("/modulo/usuarios")).Headers.Location?.OriginalString == "/Administracion/Usuarios",
    "Menú de administración abre el listado de usuarios");
var userForm = WebUtility.HtmlDecode(await securityClient.GetStringAsync("/Administracion/Usuarios/Nuevo"));
Check(userForm.Contains("Contraseña temporal") && userForm.Contains("Perfil de acceso") && userForm.Contains("data-security-copy"),
    "Alta muestra datos y controles de contraseña temporal");
var roleForm = WebUtility.HtmlDecode(await securityClient.GetStringAsync("/Administracion/Perfiles/consulta/Duplicar"));
Check(roleForm.Contains("Copia de Consulta") && roleForm.Contains("data-security-module") && roleForm.Contains("Guardar perfil"),
    "Duplicación abre un nuevo perfil con casillas de permisos");
var auditRepository = (TestSecurityAdministration)securityFactory.Services.GetRequiredService<ISecurityAdministrationRepository>();
var auditFilters = await securityClient.GetStringAsync("/Administracion/Auditoria?entity=Presupuesto&eventAction=GENERAR_IMPRESION");
Check(auditRepository.LastAuditFilter == ("Presupuesto", "GENERAR_IMPRESION") &&
    Regex.IsMatch(auditFilters, "<option value=\"GENERAR_IMPRESION\"[^>]* selected") &&
    !auditFilters.Contains("value=\"Index\""), "Auditoría conserva la acción elegida sin confundirla con la ruta MVC");
await securityClient.GetStringAsync("/Administracion/Auditoria?entity=PresupuestoRevision&action=CAMBIAR_ESTADO");
Check(auditRepository.LastAuditFilter == ("Presupuesto", "CAMBIAR_ESTADO"), "Auditoría conserva enlaces anteriores");
using (var auditRequest = new HttpRequestMessage(HttpMethod.Get, "/Administracion/Auditoria/1"))
{
    auditRequest.Headers.Add("X-Requested-With", "XMLHttpRequest");
    var audit = WebUtility.HtmlDecode(await (await securityClient.SendAsync(auditRequest)).Content.ReadAsStringAsync());
    Check(audit.Contains("Detalle de auditoría") && audit.Contains("Ana") && audit.Contains("María") &&
        !audit.Contains("secreto-fixture") && !audit.Contains("PasswordHash"), "Popup de auditoría muestra diferencias sin datos sensibles");
}
Console.WriteLine("Pruebas MVC de acceso y parciales completadas sin SQL Server.");

using var maintenanceFactory = new HytFactory(repository);
using var maintenanceClient = maintenanceFactory.CreateClient(new WebApplicationFactoryClientOptions { BaseAddress = new Uri("https://localhost"), AllowAutoRedirect = false });
async Task<string> MaintenanceToken(string path)
{
    var html = await maintenanceClient.GetStringAsync(path);
    var input = Regex.Match(html, "<input[^>]*name=\"__RequestVerificationToken\"[^>]*>").Value;
    return WebUtility.HtmlDecode(Regex.Match(input, "value=\"([^\"]+)\"").Groups[1].Value);
}
repository.Account = repository.Account with { Permissions = ["CLIENTES.CONSULTAR", "CLIENTES.EDITAR", "TARIFARIOS.CONSULTAR", "TARIFARIOS.EDITAR"] };
await maintenanceClient.PostAsync("/Acceso", new FormUrlEncodedContent(new Dictionary<string,string> { ["UserName"] = "operador", ["Password"] = repository.Password, ["__RequestVerificationToken"] = await MaintenanceToken("/Acceso") }));
var masterList = await maintenanceClient.GetStringAsync("/Maestros/clientes?search=abc&page=2");
Check(masterList.Contains("0001-A") && !masterList.Contains("<script>prueba</script>") && masterList.Contains("returnUrl=") && maintenanceFactory.Maintenance.LastPage == 2 && maintenanceFactory.Maintenance.LastSearch == "abc", "Listado real envía búsqueda/página al repositorio y conserva retorno seguro");
var clientForm = await maintenanceClient.GetStringAsync("/Maestros/clientes/Nuevo");
Check(clientForm.Contains("internal-record-code") && clientForm.Contains("__RequestVerificationToken") && clientForm.Contains("master-lookup"), "Ficha nueva MVC contiene campos, antifalsificación y popup");
var invalidMaster = await maintenanceClient.PostAsync("/Maestros/clientes/Guardar", new FormUrlEncodedContent(new Dictionary<string, string> { ["__RequestVerificationToken"] = await MaintenanceToken("/Maestros/clientes/Nuevo") }));
Check(invalidMaster.StatusCode == HttpStatusCode.OK && maintenanceFactory.Maintenance.SaveCount == 0, "Formulario inválido conserva vista y no guarda");
var savedMaster = await maintenanceClient.PostAsync("/Maestros/clientes/Guardar", new FormUrlEncodedContent(new Dictionary<string, string>
{ ["__RequestVerificationToken"] = await MaintenanceToken("/Maestros/clientes/Nuevo"), ["Fields[Codigo]"] = "0002-X", ["Fields[TipoEntidad]"] = "JURIDICA", ["Fields[RazonSocial]"] = "Cliente", ["Fields[TipoIdentificacion]"] = "RUC", ["Fields[NumeroIdentificacion]"] = "20123456789", ["Fields[Activo]"] = "true", ["ReturnUrl"] = "https://externo.example" }));
Check(savedMaster.StatusCode == HttpStatusCode.Redirect && maintenanceFactory.Maintenance.SaveCount == 1 && !savedMaster.Headers.Location!.OriginalString.Contains("externo.example"), "Guardado MVC usa repositorio y rechaza retorno externo");
Check((await maintenanceClient.GetAsync("/Maestros/clientes/999")).StatusCode == HttpStatusCode.NotFound, "Ficha inexistente responde 404");
Check((await maintenanceClient.GetAsync("/Maestros/clientes/Ayuda?page=0")).StatusCode == HttpStatusCode.BadRequest, "Ayuda rechaza página inválida");
var lookupJson = await maintenanceClient.GetStringAsync("/Maestros/clientes/Ayuda?page=2&search=tienda");
Check(lookupJson.Contains("0001-A") && maintenanceFactory.Maintenance.LastPage == 2, "Ayuda consulta resultados paginados del repositorio");
var published = await maintenanceClient.GetStringAsync("/Maestros/versiones-tarifario/2/Editar");
Check(!published.Contains("Guardar registro"), "Versión publicada abre ficha inmutable");
Check((await maintenanceClient.PostAsync("/Maestros/versiones-tarifario/Guardar", new FormUrlEncodedContent(new Dictionary<string, string> { ["Id"] = "2", ["__RequestVerificationToken"] = await MaintenanceToken("/Maestros/versiones-tarifario/2") }))).StatusCode == HttpStatusCode.BadRequest, "POST directo no modifica versión publicada");
var versionToken = await MaintenanceToken("/Maestros/versiones-tarifario/1");
Check((await maintenanceClient.PostAsync("/Maestros/versiones-tarifario/1/Publicar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = versionToken, ["rowVersion"] = Convert.ToBase64String(new byte[8]) }))).StatusCode == HttpStatusCode.BadRequest && maintenanceFactory.Maintenance.Publications == 0, "Publicación exige confirmación explícita");
var publishResponse = await maintenanceClient.PostAsync("/Maestros/versiones-tarifario/1/Publicar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = versionToken, ["rowVersion"] = Convert.ToBase64String(new byte[8]), ["confirmed"] = "true" }));
Check(publishResponse.StatusCode == HttpStatusCode.Redirect && maintenanceFactory.Maintenance.Publications == 1, "Publicación confirmada pasa versión original al repositorio");
var copyResponse = await maintenanceClient.PostAsync("/Maestros/versiones-tarifario/2/Copiar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = versionToken, ["rowVersion"] = Convert.ToBase64String(new byte[8]), ["confirmed"] = "true" }));
Check(copyResponse.StatusCode == HttpStatusCode.Redirect && copyResponse.Headers.Location!.OriginalString.StartsWith("/Maestros/versiones-tarifario/3") && maintenanceFactory.Maintenance.Copies == 1, "Copia abre la nueva versión devuelta por el repositorio");
var getPublication = await maintenanceClient.GetAsync("/Maestros/versiones-tarifario/1/Publicar");
Check(getPublication.StatusCode is HttpStatusCode.MethodNotAllowed or HttpStatusCode.NotFound && maintenanceFactory.Maintenance.Publications == 1, "Publicación no se ejecuta por GET");
repository.Account = repository.Account with { Permissions = ["CONFIGURACION.CONSULTAR", "CONFIGURACION.EDITAR", "TARIFARIOS.CONSULTAR", "CLIENTES.CONSULTAR"] };
var rateHtml = await maintenanceClient.GetStringAsync("/Configuracion/TasasIGV");
Check(rateHtml.Contains("0.18") && rateHtml.Contains("20 por página") && rateHtml.Contains("ExpectedLatest"), "Listado de vigencias muestra tasa, paginación y versión de configuración");
var invalidRate = await maintenanceClient.PostAsync("/Configuracion/TasasIGV/Agregar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = await MaintenanceToken("/Configuracion/TasasIGV"), ["From"] = "2027-01-01", ["Value"] = "18", ["Reason"] = "Cambio" }));
Check(invalidRate.StatusCode == HttpStatusCode.OK && maintenanceFactory.TaxRates.Added == 0, "Tasa 18 se rechaza; se requiere fracción como 0.18");
var staleRate = await maintenanceClient.PostAsync("/Configuracion/TasasIGV/Agregar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = await MaintenanceToken("/Configuracion/TasasIGV"), ["From"] = "2027-01-01", ["Value"] = "0.19", ["Reason"] = "Cambio", ["ExpectedLatest"] = "2025-01-01" }));
Check(staleRate.StatusCode == HttpStatusCode.OK && maintenanceFactory.TaxRates.Added == 0 && (await staleRate.Content.ReadAsStringAsync()).Contains("vigencias cambiaron"), "Vigencia concurrente se informa sin sobrescribir cambios");
var validRate = await maintenanceClient.PostAsync("/Configuracion/TasasIGV/Agregar", new FormUrlEncodedContent(new Dictionary<string,string> { ["__RequestVerificationToken"] = await MaintenanceToken("/Configuracion/TasasIGV"), ["From"] = "2027-01-01", ["Value"] = "0.19", ["Reason"] = "Cambio", ["ExpectedLatest"] = "2026-01-01" }));
Check(validRate.StatusCode == HttpStatusCode.Redirect && maintenanceFactory.TaxRates.Added == 1, "Nueva vigencia válida utiliza el repositorio");
Check((await maintenanceClient.GetStringAsync("/Maestros/versiones-tarifario/1/Predeterminados")).Contains("CodigoMonedaPredeterminada"), "Ayuda de predeterminados devuelve datos permitidos de la versión");
Check((await maintenanceClient.GetAsync("/Maestros/versiones-tarifario/2/Predeterminados")).StatusCode == HttpStatusCode.BadRequest, "Versión publicada no sirve de padre editable para partidas nuevas");
maintenanceFactory.Maintenance.Unavailable = true;
var unavailable = await maintenanceClient.GetAsync("/Maestros/clientes/1");
Check(unavailable.StatusCode == HttpStatusCode.ServiceUnavailable && !(await unavailable.Content.ReadAsStringAsync()).Contains("Datos internos de prueba"), "Fallo de datos devuelve 503 sin detalles internos");
maintenanceFactory.Maintenance.Unavailable = false;
repository.Account = repository.Account with { Permissions = MaintenanceDefinitions.All.SelectMany(entity => new[] { entity.Permission + ".CONSULTAR", entity.Permission + ".EDITAR" }).Distinct().ToArray() };
foreach (var definition in MaintenanceDefinitions.All.Where(item=>item.Key is not ("proveedor-cliente" or "tarifarios")))
{
    Check((await maintenanceClient.GetAsync($"/Maestros/{definition.Key}/Nuevo")).StatusCode == HttpStatusCode.OK && (await maintenanceClient.GetAsync($"/Maestros/{definition.Key}")).StatusCode == HttpStatusCode.OK, "Listado y ficha disponibles: " + definition.Title);
}
maintenanceFactory.Maintenance.ConcurrentSave = true;
var concurrentResponse = await maintenanceClient.PostAsync("/Maestros/clientes/Guardar", new FormUrlEncodedContent(new Dictionary<string,string>
{ ["__RequestVerificationToken"] = await MaintenanceToken("/Maestros/clientes/1/Editar"), ["Id"] = "1", ["RowVersion"] = Convert.ToBase64String(new byte[8]), ["Fields[Codigo]"] = "C1", ["Fields[TipoEntidad]"] = "JURIDICA", ["Fields[RazonSocial]"] = "Edición conservada", ["Fields[TipoIdentificacion]"] = "RUC", ["Fields[NumeroIdentificacion]"] = "20123456789", ["Fields[Activo]"] = "true" }));
var concurrentHtml = WebUtility.HtmlDecode(await concurrentResponse.Content.ReadAsStringAsync());
Check(concurrentResponse.StatusCode == HttpStatusCode.OK && concurrentHtml.Contains("registro cambió") && concurrentHtml.Contains("Edición conservada") && maintenanceFactory.Maintenance.SaveCount == 1, "Conflicto de edición conserva datos introducidos y no sobrescribe el registro");
maintenanceFactory.Maintenance.ConcurrentSave = false;
repository.Account = repository.Account with { Permissions = ["PRESUPUESTOS.CONSULTAR","PRESUPUESTOS.EDITAR"] };
var budgetListHtml=await maintenanceClient.GetStringAsync("/Presupuestos?search=cliente&from=2026-10-03&to=2026-11-03&page=2");
Check(budgetListHtml.Contains("PPTO-1")&&!budgetListHtml.Contains("Cliente <prueba>")&&maintenanceFactory.Budgets.RequestedPage==2&&maintenanceFactory.Budgets.From==new DateTimeOffset(2026,10,3,5,0,0,TimeSpan.Zero)&&maintenanceFactory.Budgets.Until==new DateTimeOffset(2026,11,4,5,0,0,TimeSpan.Zero),"Presupuestos filtra límites UTC de Perú y codifica la grilla");
var newBudgetHtml=await maintenanceClient.GetStringAsync("/Presupuestos/Nuevo");
Check(newBudgetHtml.Contains("Draft.Revision[ContratoMarco]")&&newBudgetHtml.Contains("Draft.Revision[NumeroCotizacionCliente]")&&newBudgetHtml.Contains("budget-line-template")&&newBudgetHtml.Contains("budget-lookup"),"Ficha de presupuesto incluye campos de ambos Excel y ayudas");
var budgetToken=await MaintenanceToken("/Presupuestos/Nuevo");
var invalidBudget=await maintenanceClient.PostAsync("/Presupuestos/Guardar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=budgetToken}));
Check(invalidBudget.StatusCode==HttpStatusCode.OK&&maintenanceFactory.Budgets.Saved==0,"Presupuesto inválido conserva formulario y no guarda");
var draftBudgetFields=new Dictionary<string,string>{["__RequestVerificationToken"]=budgetToken,["Draft.ClienteId"]="1",["Draft.TiendaId"]="1",["Draft.DetalleNeto"]="true",["Draft.TipoAtencion"]="EMERGENCIA",["Draft.Descripcion"]="Servicio",["Draft.Revision[FechaEmision]"]="2026-11-03",["Draft.Revision[CodigoMoneda]"]="PEN",["Draft.Revision[CodigoTipoIGV]"]="SIN_IGV",["ReturnUrl"]="https://externo.example"};
var savedBudget=await maintenanceClient.PostAsync("/Presupuestos/Guardar",new FormUrlEncodedContent(draftBudgetFields));
Check(savedBudget.StatusCode==HttpStatusCode.Redirect&&maintenanceFactory.Budgets.Saved==1&&!savedBudget.Headers.Location!.OriginalString.Contains("externo.example")&&maintenanceFactory.Budgets.LastPayload!.Contains("EMERGENCIA")&&maintenanceFactory.Budgets.LastPayload.Contains("\"RevalorizarMoneda\":true"),"Borrador de emergencia se guarda por repositorio y rechaza retorno externo");
var readBudget=await maintenanceClient.GetStringAsync("/Presupuestos/1");
Check(readBudget.Contains("<fieldset disabled")&&!readBudget.Contains("Guardar borrador")&&readBudget.Contains("budget-record-number"),"Consulta de presupuesto no permite editar campos");
Check((await maintenanceClient.GetAsync("/Presupuestos/999")).StatusCode==HttpStatusCode.NotFound,"Presupuesto inexistente responde 404");
Check((await maintenanceClient.GetAsync("/Presupuestos/2/Editar")).StatusCode==HttpStatusCode.BadRequest,"No se abre edición de una revisión enviada");
Check(WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync("/Presupuestos/2/NuevaRevision?revisionId=2")).Contains("Nueva revisión a partir"),"Corrección de revisión enviada abre nueva revisión");
var tariffBudgetJson=await maintenanceClient.GetStringAsync("/Presupuestos/Tarifario?clientId=1&date=2026-11-03&currency=USD&page=2");
Check(tariffBudgetJson.Contains("0001-A")&&tariffBudgetJson.Contains("\"price\":null")&&tariffBudgetJson.Contains("Sur, Centro"),"Ayuda de presupuesto muestra zonas informativas y USD pendiente");
var previewFields=new Dictionary<string,string>(draftBudgetFields){["Draft.Partidas[0].Cantidad"]="2.00",["Draft.Partidas[0].PrecioManual"]="10.00",["Draft.Partidas[0].CodigoMonedaOrigen"]="PEN",["Draft.Partidas[0].CodigoTipoIGVOrigen"]="SIN_IGV"};
var previewBudget=await maintenanceClient.PostAsync("/Presupuestos/Calcular",new FormUrlEncodedContent(previewFields));
var previewData=System.Text.Json.JsonDocument.Parse(await previewBudget.Content.ReadAsStringAsync());
Check(previewBudget.StatusCode==HttpStatusCode.OK&&previewData.RootElement.GetProperty("totals").GetProperty("total").GetDecimal()==23.60m,"Previsualización calcula en servidor con dos decimales");
var stateBudget=await maintenanceClient.PostAsync("/Presupuestos/Estado",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=budgetToken,["Id"]="1",["RevisionId"]="1",["HeaderVersion"]=Convert.ToBase64String(new byte[8]),["RevisionVersion"]=Convert.ToBase64String(new byte[8]),["NewState"]="ENVIADO",["EventDate"]="2026-11-03",["Responsible"]="Responsable"}));
Check(stateBudget.StatusCode==HttpStatusCode.Redirect&&maintenanceFactory.Budgets.StateChanges==1,"Acción de presupuesto envía responsable, fecha y versiones al repositorio");
var budgetEditor=await maintenanceClient.GetStringAsync("/Presupuestos/1/Editar");
Check(new[]{readBudget,budgetEditor}.All(html=>!html.Contains("Revisiones del presupuesto")&&!html.Contains("Registrar cambio de estado")&&!html.Contains("data-budget-state-form")&&!html.Contains("Crear nueva revisión")),"Consulta y edición sin apartados de revisiones ni cambio de estado");
var statesList=await maintenanceClient.GetStringAsync("/Presupuestos");
Check(statesList.Contains("Registrado")&&statesList.Contains("Aceptada")&&statesList.Contains("Rechazada")&&!statesList.Contains("Vencido")&&statesList.Contains("Cambiar estado")&&!statesList.Contains(">Registrado por<")&&!statesList.Contains(">Actualizado por<")&&!statesList.Contains(">Registro<")&&!statesList.Contains(">Actualización<"),"Listado usa estados comerciales, acción de estado y retira columnas de auditoría");
var statePopup=await maintenanceClient.GetStringAsync("/Presupuestos/2/Estado");
Check(statePopup.Contains("data-document-form")&&statePopup.Contains("Aceptada")&&!statePopup.Contains("EventDate")&&!statePopup.Contains("Responsible")&&!statePopup.Contains("VENCIDO")&&statePopup.Contains("__RequestVerificationToken"),"Popup de estado conserva CSRF y versiones sin pedir fecha ni representante");
Check((await maintenanceClient.GetAsync("/Presupuestos/88/Editar")).StatusCode==HttpStatusCode.OK,"Presupuesto observado admite edición directa");
var automaticStateFields=new Dictionary<string,string>{["__RequestVerificationToken"]=budgetToken,["Id"]="1",["RevisionId"]="1",["HeaderVersion"]=Convert.ToBase64String(new byte[8]),["RevisionVersion"]=Convert.ToBase64String(new byte[8]),["NewState"]="ENVIADO"};
var autoState=await maintenanceClient.PostAsync("/Presupuestos/Estado",new FormUrlEncodedContent(automaticStateFields));
Check(autoState.StatusCode==HttpStatusCode.Redirect&&maintenanceFactory.Budgets.StateChanges==2&&maintenanceFactory.Budgets.LastStateDate!=default&&!string.IsNullOrWhiteSpace(maintenanceFactory.Budgets.LastResponsible),"Fecha y responsable se asignan desde servidor y sesión sin campos del cliente");
foreach(var reasonState in new[]{"OBSERVADO","RECHAZADO","CANCELADO"})
{
    automaticStateFields["NewState"]=reasonState;
    Check((await maintenanceClient.PostAsync("/Presupuestos/Estado",new FormUrlEncodedContent(automaticStateFields))).StatusCode==HttpStatusCode.BadRequest,"Estado "+reasonState+" exige motivo en servidor");
}
automaticStateFields["NewState"]="VENCIDO";
Check((await maintenanceClient.PostAsync("/Presupuestos/Estado",new FormUrlEncodedContent(automaticStateFields))).StatusCode==HttpStatusCode.BadRequest,"Vencido no se puede asignar por petición manual");
repository.Account = repository.Account with { Permissions = ["PRESUPUESTOS.CONSULTAR"] };
Check((await maintenanceClient.GetAsync("/Presupuestos/Nuevo")).StatusCode==HttpStatusCode.Redirect,"Consulta de presupuestos no concede permiso de crear");
repository.Account=repository.Account with {Permissions=["CONFIGURACION.CONSULTAR","CONFIGURACION.EDITAR","PRESUPUESTOS.CONSULTAR","REPORTES.CONSULTAR"]};
var companyHtml=await maintenanceClient.GetStringAsync("/Configuracion/HYT");
Check(companyHtml.Contains("Fields[Ruc]")&&WebUtility.HtmlDecode(companyHtml).Contains("03/11/2026 13:30"),"Configuración corporativa muestra campos y fechas de Perú");
var badCompany=await maintenanceClient.PostAsync("/Configuracion/HYT/Guardar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/Configuracion/HYT"),["Fields[Ruc]"]="123",["Fields[RazonSocial]"]="HYT"}));
Check(badCompany.StatusCode==HttpStatusCode.OK&&maintenanceFactory.Company.Saves==0,"Configuración exige RUC de once dígitos");
var goodCompany=await maintenanceClient.PostAsync("/Configuracion/HYT/Guardar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/Configuracion/HYT"),["Id"]="1",["RowVersion"]=Convert.ToBase64String(new byte[8]),["Fields[Ruc]"]="00000000000",["Fields[RazonSocial]"]="HYT PRUEBA"}));
Check(goodCompany.StatusCode==HttpStatusCode.Redirect&&maintenanceFactory.Company.Saves==1,"Configuración guarda por repositorio con versión original");
Check(companyHtml.Contains("Subir firma para presupuesto")&&companyHtml.Contains("multipart/form-data"),"Configuración permite subir firma JPG/PNG");
var signatureBytes=Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6mGsAAAAASUVORK5CYII=");
async Task<HttpResponseMessage> UploadSignatureForTest(byte[] bytes,string filename)
{
    var token=await MaintenanceToken("/Configuracion/HYT");
    using var body=new MultipartFormDataContent();body.Add(new StringContent(token),"__RequestVerificationToken");body.Add(new StringContent(Convert.ToBase64String(new byte[8])),"rowVersion");body.Add(new ByteArrayContent(bytes),"signature",filename);
    return await maintenanceClient.PostAsync("/Configuracion/HYT/Firma",body);
}
await UploadSignatureForTest("<svg/>"u8.ToArray(),"firma.svg");
Check(maintenanceFactory.Company.SignatureJson is null,"Firma rechaza SVG y conserva configuración previa");
await UploadSignatureForTest(signatureBytes,"firma.png");
Check(maintenanceFactory.Company.SignatureJson is not null,"Firma PNG válida guarda metadatos privados por repositorio");
var signatureResponse=await maintenanceClient.GetAsync("/Configuracion/HYT/Firma");
Check(signatureResponse.StatusCode==HttpStatusCode.OK&&signatureResponse.Content.Headers.ContentType?.MediaType=="image/png"&&(await signatureResponse.Content.ReadAsByteArrayAsync()).SequenceEqual(signatureBytes),"Vista autorizada recupera firma con integridad");
Check(companyHtml.Contains("Fields[SubtituloCotizacion]"),"Configuración ofrece subtítulo de cotización editable");
repository.Account=repository.Account with {Permissions=repository.Account.Permissions.Concat(new[]{"PRESUPUESTOS.EDITAR"}).ToArray()};
var printHeadingView=await maintenanceClient.GetStringAsync("/Presupuestos/1");
var printHeadingEdit=await maintenanceClient.GetStringAsync("/Presupuestos/1/Editar");
foreach(var savedView in new[]{printHeadingView,printHeadingEdit})
{
    var heading=Regex.Match(savedView,@"<header class=""hyt-page-header hyt-budget-header""[\s\S]*?</header>").Value;
    Check(heading.Contains("hyt-budget-print-card")&&heading.Contains("value=\"NUEVO\"")&&heading.Contains("value=\"ANTERIOR\"")&&heading.Contains("__RequestVerificationToken")&&!savedView.Contains("Consultar impresiones conservadas"),"Consulta y edición guardadas muestran ambos formatos dentro del título sin enlace histórico");
}
Check(!(await maintenanceClient.GetStringAsync("/Presupuestos/Nuevo")).Contains("hyt-budget-print-card"),"Presupuesto sin guardar no muestra opciones de impresión");
var printToken=await MaintenanceToken("/Presupuestos/1");
foreach(var format in new[]{"NUEVO","ANTERIOR"})
{
    var generated=await maintenanceClient.PostAsync("/Presupuestos/Impresiones/Generar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=printToken,["budgetId"]="1",["revisionId"]="1",["version"]=Convert.ToBase64String(new byte[8]),["format"]=format}));
    Check(generated.StatusCode==HttpStatusCode.Redirect,"Generación HTML autorizada: "+format);
    var document=await maintenanceClient.GetAsync(generated.Headers.Location);var html=WebUtility.HtmlDecode(await document.Content.ReadAsStringAsync());
    Check(document.StatusCode==HttpStatusCode.OK&&html.StartsWith("<!DOCTYPE html>",StringComparison.OrdinalIgnoreCase)&&html.Contains("data:image/png;base64,")&&html.Contains(format=="NUEVO"?"118.00":"100.00")&&document.Headers.GetValues("Content-Security-Policy").Single().Contains("script-src 'none'"),"Reporte autónomo conserva logo, importes y política sin scripts: "+format);
    Check(html.Contains("Firma autorizada para presupuesto")&&html.Contains(Convert.ToBase64String(signatureBytes)),"Impresión conserva firma embebida desde almacenamiento: "+format);
    Check(html.Contains(GestionServicios.Application.Maintenance.CompanySettings.DefaultSubtitle),"Impresión muestra subtítulo completo desde configuración: "+format);
    Check(format=="NUEVO"?html.Contains("CENTRO SAP FICO")&&html.Contains("000SAP"):html.Contains("P. PARCIAL")&&html.Contains("COSTO FINAL SIN IGV"),"Columnas del formato "+format);
    if(format=="ANTERIOR")Check(html.Contains("Coordinador/Encargado:")&&html.Contains("Contacto de tienda"),"Formato anterior muestra contacto de tienda bajo Tienda");
    if(format=="ANTERIOR")Check(!html.Contains("class=\"report-status\"")&&!html.Contains("Coordinador / encargado")&&!html.Contains("Generado ")&&!html.Contains("Los precios unitarios"),"Formato anterior omite estado, responsables internos, generación y notas automáticas");
    Check(!(await document.Content.ReadAsStringAsync()).Contains("Cliente <prueba>"),"Reporte codifica datos de cliente: "+format);
}
Check(maintenanceFactory.Prints.Count==2&&(await maintenanceClient.GetStringAsync("/Presupuestos/Impresiones?budgetId=1&revisionId=1")).Contains("HYT-2026.10"),"Historial conserva ambas generaciones y versión del formato");
Check((await maintenanceClient.GetAsync("/Presupuestos/Impresiones/1?download=true")).Content.Headers.ContentDisposition?.DispositionType=="attachment","Impresión conservada permite descarga HTML");
var zeroPrint=await maintenanceClient.PostAsync("/Presupuestos/Impresiones/Generar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=printToken,["budgetId"]="51",["revisionId"]="51",["version"]=Convert.ToBase64String(new byte[8]),["format"]="NUEVO"}));
Check(zeroPrint.StatusCode==HttpStatusCode.Redirect&&WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync(zeroPrint.Headers.Location)).Contains("todavía no contiene partidas"),"Borrador sin partidas se identifica en la impresión");
repository.Account=repository.Account with {Permissions=["PRESUPUESTOS.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Presupuestos/Impresiones/1")).StatusCode==HttpStatusCode.Redirect,"Sin permiso de reportes no se entrega el HTML conservado");
Check(!(await maintenanceClient.GetStringAsync("/Presupuestos/1")).Contains("hyt-budget-print-card"),"Sin permiso de reportes no se ofrecen botones de impresión");
repository.Account = repository.Account with { Permissions = ["CLIENTES.CONSULTAR"] };
Check((await maintenanceClient.GetAsync("/Maestros/clientes/Nuevo")).StatusCode == HttpStatusCode.Redirect, "Sin permiso de edición no abre formulario nuevo");


repository.Account=repository.Account with {Permissions=["OC_CLIENTE.CONSULTAR","GUIAS.CONSULTAR","CONFIRMACIONES.CONSULTAR"]};
foreach(var kind in new[]{"oc-cliente","guias","confirmaciones"})
{
    var documentList=await maintenanceClient.GetStringAsync($"/Seguimiento/{kind}?search=OC&page=2&from=2026-11-01&to=2026-11-03");
    Check(documentList.Contains("20 por página")&&documentList.Contains("page=3")&&maintenanceFactory.Documents.Page==2&&maintenanceFactory.Documents.Search=="OC"&&!documentList.Contains("<script>prueba</script>"),"Listado documental filtra, pagina y codifica datos: "+kind);
    var details=WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync($"/Seguimiento/{kind}/1"));
    Check(details.Contains("03/11/2026 13:30")&&details.Contains("Fechas y responsables")&&!details.Contains("Historial de acciones")&&details.Contains("Operador"),"Ficha documental muestra responsables y fechas Perú sin historial: "+kind);
}
Check((await maintenanceClient.GetAsync("/Seguimiento/oc-cliente/999")).StatusCode==HttpStatusCode.NotFound,"Documento inexistente responde 404");
Check((await maintenanceClient.GetAsync("/Seguimiento/oc-cliente?page=0")).StatusCode==HttpStatusCode.BadRequest,"Listado documental rechaza página inválida");
var menuOc=await maintenanceClient.GetAsync("/modulo/oc-cliente");
Check(menuOc.StatusCode==HttpStatusCode.Redirect&&menuOc.Headers.Location!.ToString().Contains("/Seguimiento/oc-cliente"),"Menú OC abre el listado operativo");
repository.Account=repository.Account with {Permissions=["CLIENTES.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Seguimiento/oc-cliente")).StatusCode==HttpStatusCode.Redirect,"Consulta documental exige permiso propio");
repository.Account=repository.Account with {Permissions=["OC_CLIENTE.CONSULTAR","OC_CLIENTE.EDITAR","GUIAS.CONSULTAR","GUIAS.EDITAR","CONFIRMACIONES.CONSULTAR","CONFIRMACIONES.EDITAR","PRESUPUESTOS.CONSULTAR"]};
var guideForm=await maintenanceClient.GetStringAsync("/Seguimiento/guias/Registro?parentId=77&popup=true");
Check(guideForm.Contains("Fields[NumeroActa]")&&guideForm.Contains("Fields[FechaActa]")&&guideForm.Contains("OperationId")&&!guideForm.Contains("<!DOCTYPE html>"),"Registro de guía devuelve captura parcial para popup");
var formOperation=Guid.NewGuid();
var guideFields=new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/Seguimiento/guias/Registro?parentId=77"),["ParentId"]="77",["ParentVersion"]=Convert.ToBase64String(new byte[8]),["OperationId"]=formOperation.ToString(),["Fields[FechaActa]"]="2026-11-03",["Fields[NumeroActa]"]="GUIA-0007B",["Fields[Estado]"]="ACEPTADA",["ReturnUrl"]="https://externo.example"};
var savedGuide=await maintenanceClient.PostAsync("/Seguimiento/guias/Guardar?popup=true",new FormUrlEncodedContent(guideFields));
Check(savedGuide.StatusCode==HttpStatusCode.OK&&maintenanceFactory.Documents.Saves==1&&maintenanceFactory.Documents.LastOperation==formOperation&&!maintenanceFactory.Documents.LastPayload!.Contains("ACEPTADA")&&!(await savedGuide.Content.ReadAsStringAsync()).Contains("externo.example"),"Guía popup guarda campos permitidos con operación y retorno seguro");
guideFields["Fields[NumeroActa]"]="";
var invalidGuide=await maintenanceClient.PostAsync("/Seguimiento/guias/Guardar?popup=true",new FormUrlEncodedContent(guideFields));
Check((int)invalidGuide.StatusCode==422&&maintenanceFactory.Documents.Saves==1&&(await invalidGuide.Content.ReadAsStringAsync()).Contains(formOperation.ToString()),"Validación conserva operación y campos sin guardar");
guideFields.Remove("__RequestVerificationToken");
Check((await maintenanceClient.PostAsync("/Seguimiento/guias/Guardar",new FormUrlEncodedContent(guideFields))).StatusCode==HttpStatusCode.BadRequest,"Guardado documental sin antifalsificación se rechaza");
var editDocument=await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente/1/Editar");
Check(editDocument.Contains("MotivoCambio")&&editDocument.Contains("OriginalVersion")&&editDocument.Contains("OC-000A/01"),"Edición OC carga número externo, motivo y versión original");
var ocInformative=WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente/Registro?parentId=77&popup=true"));
Check(ocInformative.Contains("hyt-document-context")&&ocInformative.Contains("hyt-internal-code")&&ocInformative.Contains("Número de presupuesto")&&ocInformative.Contains("PPTO-1")&&ocInformative.Contains("Estado del presupuesto")&&ocInformative.Contains("Aceptada")&&!ocInformative.Contains("<select")&&!ocInformative.Contains("Revisión del presupuesto"),"OC popup muestra número en tarjeta y estado informativo sin selector de revisión");
var ocAutomaticFields=new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/Seguimiento/oc-cliente/Registro?parentId=77"),["ParentId"]="77",["ParentVersion"]=Convert.ToBase64String(new byte[8]),["OperationId"]=Guid.NewGuid().ToString(),["Fields[RevisionId]"]="999",["Fields[NumeroOCCliente]"]="OC-AUTO",["Fields[FechaEmisionOC]"]="2026-11-06",["Fields[FechaRecepcionOC]"]="2026-11-06"};
var ocAutomatic=await maintenanceClient.PostAsync("/Seguimiento/oc-cliente/Guardar?popup=true",new FormUrlEncodedContent(ocAutomaticFields));
Check(ocAutomatic.StatusCode==HttpStatusCode.OK&&System.Text.Json.JsonDocument.Parse(maintenanceFactory.Documents.LastPayload!).RootElement.GetProperty("RevisionId").GetInt64()==1,"OC usa revisión asociada por servidor y no permite cambiarla mediante campo oculto manipulado");
var ocStatePopup=WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente/1/CambiarEstado?returnUrl=%2FSeguimiento%2Foc-cliente"));
Check(ocStatePopup.Contains("Fechas y responsables")&&ocStatePopup.Contains("Registrar cambio de estado")&&ocStatePopup.Contains("name=\"originalVersion\"")&&!ocStatePopup.Contains("<!DOCTYPE html>"),"Estado OC abre partial con trazabilidad, versiones y formulario de transición");
var ownValidation=GestionServicios.Application.Followup.DocumentFields.Validate(GestionServicios.Application.Followup.DocumentKind.Resolve("oc-cliente"),new Dictionary<string,string?>{["RevisionId"]="1",["NumeroOCCliente"]="000A",["FechaEmisionOC"]="2026-11-03",["FechaRecepcionOC"]="2026-11-03",["Importe"]="1.001"},null,false);
Check(ownValidation.Issues.Any(issue=>issue.Field=="Importe"),"Captura OC rechaza más de dos decimales sin redondear silenciosamente");
var popupList=await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente");
Check(!popupList.Contains("hyt-action-registered"),"OC sin confirmación ni marcado conserva botones sin borde de registro");
Check(popupList.Contains("data-document-open")&&WebUtility.HtmlDecode(popupList).Contains("/Seguimiento/confirmaciones/Registro?"),"Listado OC ofrece confirmación del cliente mediante popup");
var budgetWithGuide=await maintenanceClient.GetStringAsync("/Presupuestos");
Check(budgetWithGuide.Contains("data-document-open")&&WebUtility.HtmlDecode(budgetWithGuide).Contains("/Seguimiento/guias/Registro?"),"Listado presupuesto ofrece guía mediante popup");
var registeredDocumentPermissions=repository.Account.Permissions;
repository.Account=repository.Account with{Permissions=registeredDocumentPermissions.Concat(["EJECUCION.CONSULTAR"]).ToArray()};
maintenanceFactory.Budgets.RegisteredDocuments=true;
var registeredDocuments=WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync("/Presupuestos"));
Check(Regex.Matches(registeredDocuments,"hyt-action-registered").Count==3&&registeredDocuments.Contains("/Seguimiento/guias/1")&&registeredDocuments.Contains("/Seguimiento/oc-cliente/1")&&registeredDocuments.Contains("/Seguimiento/ejecucion/1"),"Presupuesto resalta tres documentos registrados y enlaza fichas con detalle y adjuntos");
Check(registeredDocuments.Contains("/Seguimiento/oc-cliente/1/Editar?")&&registeredDocuments.Contains("popup=True",StringComparison.OrdinalIgnoreCase),"OC registrada desde presupuesto reutiliza edición popup y conserva señal de registro");
maintenanceFactory.Budgets.RegisteredDocuments=false;
repository.Account=repository.Account with{Permissions=registeredDocumentPermissions};
repository.Account=repository.Account with {Permissions=["GUIAS.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Seguimiento/guias/Registro?parentId=77")).StatusCode==HttpStatusCode.Redirect,"Consulta de guías no concede permiso para registrar");
var storageRoot=Path.Combine(Path.GetTempPath(),"hyt-file-tests-"+Guid.NewGuid().ToString("N"));
try
{
    var storageConfiguration=new Microsoft.Extensions.Configuration.ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string,string?>{["DocumentStorage:Root"]=storageRoot}).Build();
    var privateFiles=new GestionServicios.Web.Services.PrivateDocumentFileStore(storageConfiguration,maintenanceFactory.Services.GetRequiredService<IWebHostEnvironment>());
    var kindGuide=GestionServicios.Application.Followup.DocumentKind.Resolve("guias");
    var fileOperation=Guid.NewGuid();
    var pdfBytes=System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj<</Type/Catalog>>endobj\n%%EOF");
    Microsoft.AspNetCore.Http.FormFile Upload(byte[] data,string name)=>new(new MemoryStream(data),0,data.Length,"Attachment",name);
    var metadata=await privateFiles.SaveAsync(Upload(pdfBytes,"C:\\escaneados\\Guía 0007B.pdf"),kindGuide,fileOperation,CancellationToken.None);
    Check(metadata.NombreOriginal=="Guía 0007B.pdf"&&metadata.ClaveAlmacenamiento==fileOperation.ToString("N")+".pdf"&&metadata.TipoContenido=="application/pdf","Archivo privado usa nombre limpio, clave opaca y tipo verificado");
    Check((await privateFiles.ReadAsync(metadata,CancellationToken.None))!.SequenceEqual(pdfBytes),"Archivo recuperado conserva contenido y huella");
    await privateFiles.SaveAsync(Upload(pdfBytes,"Guía 0007B.pdf"),kindGuide,fileOperation,CancellationToken.None);
    Check(Directory.GetFiles(storageRoot).Length==1,"Reintento de archivo no crea duplicados ni deja temporales");
    async Task RejectedFile(Func<Task> action,string label)
    {
        try{await action();}catch(MaintenanceException){Check(true,label);return;}throw new InvalidOperationException(label);
    }
    await RejectedFile(()=>privateFiles.SaveAsync(Upload(System.Text.Encoding.ASCII.GetBytes("archivo ejecutable"),"falso.pdf"),kindGuide,Guid.NewGuid(),CancellationToken.None),"Extensión PDF no permite contenido de otro formato");
    await RejectedFile(()=>privateFiles.SaveAsync(Upload(new byte[]{137,80,78,71,13,10,26,10},"guia.png"),kindGuide,Guid.NewGuid(),CancellationToken.None),"Guía exige escaneado PDF aunque la imagen sea válida");
    await RejectedFile(()=>privateFiles.SaveAsync(Upload(System.Text.Encoding.ASCII.GetBytes("%PDF-2.0 diferente"),"Guía.pdf"),kindGuide,fileOperation,CancellationToken.None),"Misma operación no puede reemplazar contenido privado previo");
    await File.WriteAllTextAsync(Path.Combine(storageRoot,metadata.ClaveAlmacenamiento),new string('x',pdfBytes.Length));
    await RejectedFile(()=>privateFiles.ReadAsync(metadata,CancellationToken.None),"Archivo físico alterado se rechaza por integridad");
    await RejectedFile(()=>privateFiles.ReadAsync(metadata with {ClaveAlmacenamiento="../fuera.pdf"},CancellationToken.None),"Referencia de archivo impide salir de la carpeta privada");
}
finally
{
    var resolvedStorage=Path.GetFullPath(storageRoot);var expectedParent=Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
    if(Path.GetDirectoryName(resolvedStorage)!=expectedParent||!Regex.IsMatch(Path.GetFileName(resolvedStorage),@"^hyt-file-tests-[a-f0-9]{32}$"))throw new InvalidOperationException("Ruta de limpieza de pruebas no válida.");
    if(Directory.Exists(resolvedStorage))Directory.Delete(resolvedStorage,true);
}
repository.Account=repository.Account with {Permissions=["GUIAS.CONSULTAR","GUIAS.EDITAR"]};
using(var multipart=new MultipartFormDataContent())
{
    multipart.Add(new StringContent(await MaintenanceToken("/Seguimiento/guias/Registro?parentId=77")),"__RequestVerificationToken");
    multipart.Add(new StringContent("77"),"ParentId");multipart.Add(new StringContent(Convert.ToBase64String(new byte[8])),"ParentVersion");
    multipart.Add(new StringContent(Guid.NewGuid().ToString()),"OperationId");multipart.Add(new StringContent("2026-11-03"),"Fields[FechaActa]");multipart.Add(new StringContent("GUIA-0007B"),"Fields[NumeroActa]");
    multipart.Add(new ByteArrayContent(System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%%EOF")),"Attachment","guia.pdf");
    var uploadResponse=await maintenanceClient.PostAsync("/Seguimiento/guias/Guardar?popup=true",multipart);
    Check(uploadResponse.StatusCode==HttpStatusCode.OK&&maintenanceFactory.Documents.LastFile is not null,"Carga multipart conserva archivo privado y envía metadatos al repositorio");
}
var recoveredFile=await maintenanceClient.GetAsync("/Seguimiento/guias/Archivos/100");
Check(recoveredFile.StatusCode==HttpStatusCode.OK&&recoveredFile.Content.Headers.ContentType?.MediaType=="application/pdf"&&(await recoveredFile.Content.ReadAsStringAsync()).StartsWith("%PDF-")&&recoveredFile.Headers.GetValues("X-Content-Type-Options").Single()=="nosniff","Visualización de adjunto requiere consulta y entrega tipo seguro");
var downloadFile=await maintenanceClient.GetAsync("/Seguimiento/guias/Archivos/100?download=true");
Check(downloadFile.Content.Headers.ContentDisposition?.DispositionType=="attachment","Adjunto privado permite descarga con nombre original");
repository.Account=repository.Account with {Permissions=["CLIENTES.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Seguimiento/guias/Archivos/100")).StatusCode==HttpStatusCode.Redirect,"Archivo privado no se abre sin permiso del documento");
repository.Account=repository.Account with {Permissions=["GUIAS.CONSULTAR","GUIAS.EDITAR"]};
var statePage=await maintenanceClient.GetStringAsync("/Seguimiento/guias/1");
Check(!statePage.Contains("Registrar cambio de estado")&&!statePage.Contains("name=\"eventDate\"")&&!statePage.Contains("Guardar registro")&&statePage.Contains("<fieldset")&&statePage.Contains("NumeroActa"),"Consulta guía usa mismo formulario bloqueado sin flujo de estados");
var stateFields=new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/Seguimiento/guias/1"),["parentId"]="1",["parentVersion"]=Convert.ToBase64String(new byte[8]),["originalVersion"]=Convert.ToBase64String(new byte[8]),["operationId"]=Guid.NewGuid().ToString(),["state"]="ACEPTADA",["eventDate"]="2026-11-04",["responsible"]="Representante del cliente"};
Check((await maintenanceClient.PostAsync("/Seguimiento/guias/1/Estado",new FormUrlEncodedContent(stateFields))).StatusCode==HttpStatusCode.BadRequest&&maintenanceFactory.Documents.StateChanges==0,"Guía rechaza cambios de estado incluso con permiso de edición");
repository.Account=repository.Account with {Permissions=["GUIAS.CONSULTAR"]};
Check((await maintenanceClient.PostAsync("/Seguimiento/guias/1/Estado",new FormUrlEncodedContent(stateFields))).StatusCode==HttpStatusCode.Redirect,"Consulta documental no concede cambios de estado");
repository.Account=repository.Account with {Permissions=["PRESUPUESTOS.CONSULTAR","PRESUPUESTOS.EDITAR","OC_CLIENTE.CONSULTAR"]};
var emergencyDetails=await maintenanceClient.GetStringAsync("/Presupuestos/55");
Check(!emergencyDetails.Contains("name=\"NewState\"")&&!emergencyDetails.Contains("Revisiones del presupuesto")&&!emergencyDetails.Contains("Registrar cambio de estado"),"Consulta presupuesto oculta revisiones y reserva estado al listado");
var emergencyWithOrder=await maintenanceClient.GetStringAsync("/Presupuestos/55/Estado");
Check(emergencyWithOrder.Contains("name=\"NewState\"")&&emergencyWithOrder.Contains("value=\"ENVIADO\"")&&!emergencyWithOrder.Contains("value=\"CANCELADO\"")&&!emergencyWithOrder.Contains("Editar borrador"),"Popup del listado permite regularizar emergencia sin editar ni cancelar presupuesto");
var emergencySent=await maintenanceClient.GetStringAsync("/Presupuestos/56/Estado");
Check(emergencySent.Contains("value=\"APROBADO\"")&&!emergencySent.Contains("value=\"RECHAZADO\""),"Regularización enviada con OC ofrece solo aprobación");
Check(!(await maintenanceClient.GetStringAsync("/Presupuestos/57")).Contains("name=\"NewState\""),"Invitación con OC no ofrece regularización de emergencia");
var confirmedList=await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente?confirmed=true");
Check(confirmedList.Contains("CONF-0001A")&&confirmedList.Contains("236.00")&&confirmedList.Contains("name=\"confirmed\"")&&maintenanceFactory.Documents.LastConfirmed==true,"OC muestra confirmación, importe y filtro de confirmadas");
Check((await maintenanceClient.GetAsync("/Seguimiento/oc-cliente?confirmed=incorrecto")).StatusCode==HttpStatusCode.BadRequest,"Filtro de confirmación rechaza valores inválidos");
var turnstileHandler=new StubCloudflareHandler();using var turnstileHttp=new HttpClient(turnstileHandler);
var turnstileSettings=new GestionServicios.Web.Security.CloudflareTurnstileSettings{Enabled=true,SiteKey="sitekey-prueba",SecretKey="secret-prueba",AllowedHostnames=["hyt.test"]};
var testHost=new TestHostEnvironment();
var turnstileService=new GestionServicios.Web.Security.CloudflareTurnstileVerifier(turnstileHttp,Microsoft.Extensions.Options.Options.Create(turnstileSettings),testHost,TimeProvider.System);
string Challenge(bool success=true,string action="login",string hostname="hyt.test",DateTimeOffset? created=null)=>System.Text.Json.JsonSerializer.Serialize(new{success,action,hostname,challenge_ts=created??DateTimeOffset.UtcNow});
turnstileHandler.Payload=Challenge();
Check(await turnstileService.VerifyAsync("token-prueba",CancellationToken.None)&&turnstileHandler.Uri?.AbsoluteUri=="https://challenges.cloudflare.com/turnstile/v0/siteverify","Turnstile valida token con Siteverify, acción, dominio y fecha");
var requestsBefore=turnstileHandler.Requests;
Check(!await turnstileService.VerifyAsync(null,CancellationToken.None)&&!await turnstileService.VerifyAsync(new string('x',2049),CancellationToken.None)&&turnstileHandler.Requests==requestsBefore,"Turnstile rechaza token ausente o demasiado largo antes de llamar al servicio");
turnstileHandler.Payload=Challenge(action:"otro-formulario");Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Turnstile rechaza token emitido para otra acción");
turnstileHandler.Payload=Challenge(hostname:"ajeno.test");Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Turnstile rechaza dominio no autorizado");
turnstileHandler.Payload=Challenge(created:DateTimeOffset.UtcNow.AddMinutes(-6));Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Turnstile rechaza desafío vencido");
turnstileHandler.Payload=Challenge(success:false);Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Turnstile respeta rechazo del servicio");
turnstileHandler.Payload="JSON no válido";Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Respuesta inválida de Turnstile no permite acceso");
turnstileHandler.Fail=true;Check(!await turnstileService.VerifyAsync("token",CancellationToken.None),"Fallo de red de Turnstile no omite la verificación");turnstileHandler.Fail=false;
turnstileSettings.Enabled=false;testHost.EnvironmentName="Production";
Check(turnstileService.Required&&!await turnstileService.VerifyAsync(null,CancellationToken.None),"Turnstile es obligatorio en producción aunque se intente desactivar");
var fakeR2=new FakeR2Objects();var r2Store=new GestionServicios.Web.Services.R2DocumentFileStore(fakeR2);var r2Kind=GestionServicios.Application.Followup.DocumentKind.Resolve("guias");
var r2Bytes=System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%%EOF");var r2Operation=Guid.NewGuid();
Microsoft.AspNetCore.Http.FormFile R2Upload(byte[] data)=>new(new MemoryStream(data),0,data.Length,"Attachment","guia.pdf");
var r2Metadata=await r2Store.SaveAsync(R2Upload(r2Bytes),r2Kind,r2Operation,CancellationToken.None);
await r2Store.SaveAsync(R2Upload(r2Bytes),r2Kind,r2Operation,CancellationToken.None);
Check(fakeR2.Files.Count==1&&(await r2Store.ReadAsync(r2Metadata,CancellationToken.None))!.SequenceEqual(r2Bytes),"R2 conserva contenido, verifica huella y reutiliza reintento sin sobrescribir");
var rejectedR2=false;try{await r2Store.SaveAsync(R2Upload(System.Text.Encoding.ASCII.GetBytes("%PDF-2.0 alterado")),r2Kind,r2Operation,CancellationToken.None);}catch(MaintenanceException){rejectedR2=true;}
Check(rejectedR2&&fakeR2.Files.Count==1,"R2 no reemplaza un archivo existente con contenido diferente");
fakeR2.Files[r2Metadata.ClaveAlmacenamiento][0]=0;
var corruptR2=false;try{await r2Store.ReadAsync(r2Metadata,CancellationToken.None);}catch(MaintenanceException){corruptR2=true;}
Check(corruptR2,"R2 rechaza contenido alterado antes de entregarlo al usuario");
foreach(var code in new[]{"OC","GUIA","CONFIRMACION","PRESUPUESTO"})
{
    var scopedKind=new GestionServicios.Application.Followup.DocumentKind("test",code,"Test","TEST");
    var scopedFolder=GestionServicios.Web.Services.DocumentUploadRules.Folder(code,"DOC 001/A");
    var scopedOperation=Guid.NewGuid();
    var scoped=await r2Store.SaveAsync(R2Upload(r2Bytes),scopedKind,scopedOperation,CancellationToken.None,scopedFolder);
    await r2Store.SaveAsync(R2Upload(r2Bytes),scopedKind,scopedOperation,CancellationToken.None,scopedFolder);
    Check(scoped.ClaveAlmacenamiento==scopedFolder+"/"+scopedOperation.ToString("N")+".pdf"&&(await r2Store.ReadAsync(scoped,CancellationToken.None))!.SequenceEqual(r2Bytes),"R2 organiza carpeta por número y conserva reintento/lectura: "+code);
}
foreach(var invalid in new[]{"../"+r2Operation.ToString("N")+".pdf","OCCliente/../"+r2Operation.ToString("N")+".pdf","Guia/A/B/"+r2Operation.ToString("N")+".pdf",r2Operation.ToString("N")+".pdf\n"})
{
    var rejectedPath=false;try{GestionServicios.Web.Services.DocumentUploadRules.Reference(invalid);}catch(MaintenanceException){rejectedPath=true;}
    Check(rejectedPath,"Referencia de archivo rechaza ruta fuera de estructura permitida");
}
Check(GestionServicios.Web.Services.DocumentUploadRules.Folder("FIRMA","")=="ConfigHYT/Firma","Firma usa carpeta exclusiva ConfigHYT/Firma");
Check(new GestionServicios.Web.Services.CloudflareR2Settings().BucketName=="hyt-documentos","Nombre predeterminado de bucket HYT coincide con el entorno de ejemplo");
await MaintenanceRevisionChecks.Run(Check);
await CaptchaChecks.Run(Check);
var cloudflareAccess=new TestAccessRepository();
using(var cloudflareFactory=new HytFactory(cloudflareAccess,true))
using(var cloudflareClient=cloudflareFactory.CreateClient(new WebApplicationFactoryClientOptions{BaseAddress=new Uri("https://localhost"),AllowAutoRedirect=false}))
{
    var loginWithWidget=await cloudflareClient.GetStringAsync("/Acceso");
    Check(loginWithWidget.Contains("cf-turnstile")&&loginWithWidget.Contains("data-action=\"login\"")&&!loginWithWidget.Contains("secret-prueba"),"Acceso renderiza widget público sin exponer la clave secreta");
    async Task<HttpResponseMessage> CloudflareLogin(string token)
    {
        var html=await cloudflareClient.GetStringAsync("/Acceso");var input=Regex.Match(html,"<input[^>]*name=\"__RequestVerificationToken\"[^>]*>").Value;
        var csrf=WebUtility.HtmlDecode(Regex.Match(input,"value=\"([^\"]+)\"").Groups[1].Value);
        return await cloudflareClient.PostAsync("/Acceso",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=csrf,["UserName"]="operador",["Password"]=cloudflareAccess.Password,["cf-turnstile-response"]=token}));
    }
    var deniedChallenge=await CloudflareLogin("token-rechazado");
    Check(deniedChallenge.StatusCode==HttpStatusCode.OK&&cloudflareAccess.NameLookups==0,"Token Turnstile rechazado impide consultar y autenticar al usuario");
    cloudflareFactory.Turnstile.Payload=Challenge(hostname:"localhost");
    var verifiedLogin=await CloudflareLogin("token-valido");
    Check(verifiedLogin.StatusCode==HttpStatusCode.Redirect&&cloudflareAccess.NameLookups==1,"Acceso valida Turnstile en servidor antes de las credenciales");
}
repository.Account=repository.Account with {Permissions=["GUIAS.CONSULTAR","GUIAS.EDITAR","PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR","EJECUCION.CONSULTAR","CONFIRMACIONES.CONSULTAR"]};
var guideWithEvidence=await maintenanceClient.GetStringAsync("/Seguimiento/guias/1");
Check(guideWithEvidence.Contains("guia.pdf")&&guideWithEvidence.Contains("/Seguimiento/guias/Archivos/100")&&guideWithEvidence.Contains("Descargar"),"Consulta muestra enlaces para recuperar el adjunto registrado");
Check(guideWithEvidence.Contains("/Seguimiento/oc-cliente/1")&&guideWithEvidence.Contains("/Seguimiento/ejecucion/1")&&guideWithEvidence.Contains("/Presupuestos/1"),"Ficha enlaza documentos del expediente con permiso de consulta");
var contextRead=await maintenanceClient.GetAsync("/Seguimiento/ejecucion/Expediente?parentId=1");
Check(contextRead.StatusCode==HttpStatusCode.Redirect&&contextRead.Headers.Location!.ToString().Contains("/Seguimiento/ejecucion/1"),"Consulta del expediente permite abrir ejecución sin permiso de editar");
Check((await maintenanceClient.GetAsync("/Seguimiento/ejecucion/Expediente?parentId=77")).StatusCode==HttpStatusCode.NotFound,"Documento aún no registrado no abre una ficha vacía");
var guidePopupRead=await maintenanceClient.GetStringAsync("/Seguimiento/guias/1?popup=true&returnUrl=/Presupuestos");
Check(!guidePopupRead.Contains("<!DOCTYPE html>")&&guidePopupRead.Contains("Editar registro")&&!guidePopupRead.Contains("Guardar registro")&&guidePopupRead.Contains("disabled=\"disabled\""),"Guía consulta en popup permite abrir edición y mantiene campos bloqueados");
var executionPopupRead=await maintenanceClient.GetStringAsync("/Seguimiento/ejecucion/1?popup=true");
Check(executionPopupRead.Contains("data-document-kind=\"EJECUCION\"")&&!executionPopupRead.Contains("Cambiar estado")&&!executionPopupRead.Contains("name=\"state\""),"Consulta ejecución reutiliza ficha de registro sin cambio de estado");
var guideEditNavigation=await maintenanceClient.GetStringAsync("/Seguimiento/guias/1/Editar?popup=true");
Check(!guideEditNavigation.Contains("Ver registro y estados")&&guideEditNavigation.Contains("Guardar registro")&&guideEditNavigation.Contains("Cancelar"),"Edición documental conserva guardar/cancelar sin enlace a registro y estados");
repository.Account=repository.Account with {Permissions=["OC_CLIENTE.CONSULTAR","FACTURADO.CONSULTAR","FACTURADO.EDITAR"]};
var markForm=await maintenanceClient.GetStringAsync("/OCCliente/1/Facturado?popup=true");
Check(markForm.Contains("name=\"Series\"")&&markForm.Contains("Boleta (03)")&&!markForm.Contains("<!DOCTYPE html>"),"Facturado abre captura popup con tipos 01/03");
var markFields=new Dictionary<string,string>{["__RequestVerificationToken"]=await MaintenanceToken("/OCCliente/1/Facturado"),["OrderId"]="1",["OrderVersion"]=Convert.ToBase64String(new byte[8]),["OperationId"]=Guid.NewGuid().ToString(),["Date"]="2026-11-03",["Type"]="01",["Series"]="F001",["Number"]="000012A",["ReturnUrl"]="https://ajeno.example"};
var marked=await maintenanceClient.PostAsync("/OCCliente/1/Facturado?popup=true",new FormUrlEncodedContent(markFields));
Check(marked.StatusCode==HttpStatusCode.OK&&maintenanceFactory.InvoiceMarks.Saves==1&&maintenanceFactory.InvoiceMarks.Last?.Number=="000012A"&&!(await marked.Content.ReadAsStringAsync()).Contains("ajeno.example"),"Facturado MVC conserva número externo y retorno seguro");
markFields["Number"]="";
Check((int)(await maintenanceClient.PostAsync("/OCCliente/1/Facturado?popup=true",new FormUrlEncodedContent(markFields))).StatusCode==422&&maintenanceFactory.InvoiceMarks.Saves==1,"Facturado incompleto no llega al repositorio");
Check((await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente")).Contains("/OCCliente/1/Facturado"),"OC ofrece botón Facturado según permiso");
maintenanceFactory.Documents.RichList=true;
var compactOc=WebUtility.HtmlDecode(await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente"));
Check(!compactOc.Contains("<th scope=\"col\">Facturado</th>")&&compactOc.Contains("hyt-action-registered")&&!compactOc.Contains("<th scope=\"col\">Adjuntos de guía</th>")&&compactOc.Contains("<th scope=\"col\">Guía</th>"),"OC sin columna Facturado resalta el botón cuando existe comprobante vigente");
maintenanceFactory.InvoiceMarks.HasConfirmation=false;
var noConfirmationInvoice=await maintenanceClient.GetStringAsync("/OCCliente/2/Facturado?popup=true");
Check(noConfirmationInvoice.Contains("Registra primero la confirmación del cliente")&&noConfirmationInvoice.Contains("disabled=\"disabled\"")&&!noConfirmationInvoice.Contains("Guardar referencia"),"Formulario Facturado bloqueado sin confirmación vigente de cliente");
maintenanceFactory.InvoiceMarks.HasConfirmation=true;
var receiptSummary=await maintenanceClient.GetStringAsync("/OCCliente/1/Facturado/Consulta");
Check(receiptSummary.Contains("F001")&&receiptSummary.Contains("000012A")&&!receiptSummary.Contains("<form"),"Consulta de comprobante presenta datos sin edición");
repository.Account=repository.Account with {Permissions=["OC_CLIENTE.CONSULTAR","FACTURADO.CONSULTAR"]};
Check((await maintenanceClient.PostAsync("/OCCliente/1/Facturado",new FormUrlEncodedContent(markFields))).StatusCode==HttpStatusCode.Redirect,"Consulta de Facturado no concede edición");
Check((await maintenanceClient.GetAsync("/Seguimiento/guias/1/Adjuntos")).StatusCode==HttpStatusCode.Redirect,"Adjuntos de guía exigen su permiso de consulta");

repository.Account=repository.Account with {Permissions=["REPORTES.CONSULTAR","PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR"]};
foreach(var reportKind in GestionServicios.Application.Reporting.ReportKinds.All.Keys)
{
    var reportHtml=await maintenanceClient.GetStringAsync($"/Reportes?kind={reportKind}&page=2");
    Check(reportHtml.Contains("236.00")&&reportHtml.Contains("20 por página")&&reportHtml.Contains("page=3")&&!reportHtml.Contains("<script>prueba</script>")&&maintenanceFactory.Reporting.Page==2,"Reporte HTML pagina, muestra moneda y codifica datos: "+reportKind);
}
Check((await maintenanceClient.GetAsync("/Reportes?kind=INEXISTENTE")).StatusCode==HttpStatusCode.BadRequest,"Catálogo de reportes rechaza tipos arbitrarios");
repository.Account=repository.Account with {Permissions=["REPORTES.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Reportes?kind=PRESUPUESTO")).StatusCode==HttpStatusCode.Redirect,"Reporte exige además permiso del documento fuente");
repository.Account=repository.Account with {Permissions=["REPORTES.CONSULTAR","PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR"]};
var dashboardHtml=await maintenanceClient.GetStringAsync("/");
Check(Regex.IsMatch(dashboardHtml,@"href=""/Reportes(?:\?[^""]*)?""[^>]*>Consultar reportes</a>")&&!Regex.IsMatch(dashboardHtml,@"href=""/Reportes/Imprimir[^""]*""[^>]*>Consultar reportes</a>"),"Acceso Consultar reportes del dashboard genera ruta de filtros, no impresión");
var reportsNavigation=await maintenanceClient.GetAsync("/modulo/reportes");
Check(reportsNavigation.StatusCode==HttpStatusCode.Redirect&&reportsNavigation.Headers.Location?.OriginalString=="/Reportes","Menú Reportes redirige a consulta y nunca a impresión");
Check(dashboardHtml.Contains("Resumen de gestión")&&dashboardHtml.Contains("hyt-dashboard-data")&&dashboardHtml.Contains("36,000.00")&&dashboardHtml.Contains("USD")&&dashboardHtml.Contains("PEN"),"Inicio conecta indicadores y separa importes por moneda");
Check(dashboardHtml.Contains("data-chart=\"trend\"")&&dashboardHtml.Contains("data-chart=\"states\"")&&dashboardHtml.Contains("data-chart=\"donut\"")&&dashboardHtml.Contains("Consultar cifras de los gráficos"),"Dashboard ofrece curvas, columnas, dona y detalle accesible");
Check((await maintenanceClient.GetAsync("/?from=2026-11-10&to=2026-11-01")).StatusCode==HttpStatusCode.BadRequest,"Dashboard rechaza rango de fechas invertido");
Check(dashboardHtml.Contains("data-chart-group")&&dashboardHtml.Contains("data-chart-series=\"Confirmaciones\"")&&dashboardHtml.Contains("data-chart-series=\"Marcados\"")&&!dashboardHtml.Contains("Primera confirmación")&&!dashboardHtml.Contains("Primer marcado")&&dashboardHtml.Contains("Expedientes por completar"),"Panel muestra Confirmación del cliente y Facturado sin denominaciones de primeros eventos");
Check((await maintenanceClient.GetAsync("/?queuePage=0")).StatusCode==HttpStatusCode.BadRequest&&(await maintenanceClient.GetAsync("/?client="+new string('x',201))).StatusCode==HttpStatusCode.BadRequest,"Dashboard valida página y longitud de filtros comerciales");
var observedBudgetHtml=await maintenanceClient.GetStringAsync("/Presupuestos?state=OBSERVADO");
Check(observedBudgetHtml.Contains("value=\"OBSERVADO\" selected")&&observedBudgetHtml.Contains("Rechazada")&&!observedBudgetHtml.Contains("Vencido"),"Filtro presupuesto conserva todos los estados válidos del dashboard");
Check((await maintenanceClient.GetAsync("/Presupuestos?state=NO_EXISTE")).StatusCode==HttpStatusCode.BadRequest&&(await maintenanceClient.GetAsync("/Presupuestos?from=no-fecha")).StatusCode==HttpStatusCode.BadRequest,"Listado presupuesto rechaza estado inexistente y fecha inválida");
repository.Account=repository.Account with {Permissions=["PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR","GUIAS.CONSULTAR","FACTURADO.CONSULTAR"]};
var kpiBudget=await maintenanceClient.GetStringAsync("/Presupuestos?page=2&type=EMERGENCIA&pendingGuide=true&pendingOrder=true");
Check(Regex.Matches(kpiBudget,"data-kpi-state=").Count==6&&kpiBudget.Contains(">41</strong>")&&!kpiBudget.Contains("Cliente (nombre)")&&!kpiBudget.Contains("Tienda (nombre)"),"KPI presupuesto incluye seis estados y total completo sin filtros comerciales separados");
Check(maintenanceFactory.Budgets.LastType=="EMERGENCIA"&&maintenanceFactory.Budgets.PendingGuide&&maintenanceFactory.Budgets.PendingOrder&&kpiBudget.Contains("pendingGuide=true")&&kpiBudget.Contains("pendingOrder=true"),"Filtros presupuesto llegan al repositorio y persisten al paginar");
Check((await maintenanceClient.GetAsync("/Presupuestos?type=OTRO")).StatusCode==HttpStatusCode.BadRequest&&(await maintenanceClient.GetAsync("/Presupuestos?pendingGuide=incorrecto")).StatusCode==HttpStatusCode.BadRequest,"Presupuesto rechaza tipo y check inválidos");
var kpiOrder=await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente?page=2&pendingConfirmation=true&pendingInvoice=true");
Check(Regex.Matches(kpiOrder,"data-kpi-state=").Count==4&&kpiOrder.Contains(">41</strong>")&&maintenanceFactory.Documents.PendingConfirmation&&maintenanceFactory.Documents.PendingInvoice&&!kpiOrder.Contains("Cliente (nombre)")&&!kpiOrder.Contains("Tienda (nombre)"),"OC muestra cuatro KPI y aplica checks de confirmación y facturación");
var kpiGuide=await maintenanceClient.GetStringAsync("/Seguimiento/guias");
Check(!kpiGuide.Contains("Cliente (nombre)")&&!kpiGuide.Contains("Tienda (nombre)")&&kpiGuide.Contains("/Seguimiento/guias/1/Adjuntos"),"Guía simplifica filtros y permite abrir adjuntos desde listado");
repository.Account=repository.Account with {Permissions=["OC_CLIENTE.CONSULTAR"]};
Check(!(await maintenanceClient.GetStringAsync("/Seguimiento/oc-cliente")).Contains("name=\"pendingInvoice\"")&&(await maintenanceClient.GetAsync("/Seguimiento/oc-cliente?pendingInvoice=true")).StatusCode==HttpStatusCode.Redirect,"Pendientes de facturar respeta permisos de consulta");
repository.Account=repository.Account with {Permissions=["REPORTES.CONSULTAR","PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR","GUIAS.CONSULTAR","CONFIRMACIONES.CONSULTAR","FACTURADO.CONSULTAR"]};
var integralReport=await maintenanceClient.GetStringAsync("/Reportes?kind=PRESUPUESTO&state=APROBADO&type=EMERGENCIA&pendingGuide=true&pendingOrder=true&search=Cliente&page=2");
Check(!integralReport.Contains("Cliente (nombre)")&&!integralReport.Contains("Tienda (nombre)")&&!integralReport.Contains("Seguimiento de OC")&&integralReport.Contains("name=\"state\"")&&integralReport.Contains("CONF-001")&&integralReport.Contains("F001-000012A")&&integralReport.Contains("Registro:"),"Reporte integral elimina filtros antiguos e incluye referencias y registro de relaciones");
Check(maintenanceFactory.Reporting.Filters is {State:"APROBADO",Type:"EMERGENCIA",PendingGuide:true,PendingOrder:true,Search:"Cliente"}&&integralReport.Contains("page=3")&&integralReport.Contains("pendingOrder=true"),"Filtros y pendientes de reporte se conservan en repositorio y paginación");
var integralPrint=await maintenanceClient.GetStringAsync("/Reportes/Imprimir?kind=OC_CLIENTE&state=TRABAJO_COMPLETADO&pendingInvoice=true");
Check(maintenanceFactory.Reporting.Filters?.Export==true&&Regex.Matches(integralPrint,"<tr>").Count==42&&integralPrint.Contains("/images/logo-hyt.png")&&!integralPrint.Contains("hyt-sidebar")&&integralPrint.Contains("Pendiente de facturar")&&integralPrint.Contains("Completado"),"HTML completo imprime más de 20 filas con logo, estado y filtros aplicados");
Check((await maintenanceClient.GetAsync("/Reportes?kind=OC_CLIENTE&state=APROBADO")).StatusCode==HttpStatusCode.BadRequest&&(await maintenanceClient.GetAsync("/Reportes?kind=PRESUPUESTO&type=OTRO")).StatusCode==HttpStatusCode.BadRequest,"Reporte valida estados de cada fuente y tipo de atención");
repository.Account=repository.Account with {Permissions=["REPORTES.CONSULTAR","OC_CLIENTE.CONSULTAR"]};
Check((await maintenanceClient.GetAsync("/Reportes/Imprimir?kind=OC_CLIENTE&pendingInvoice=true")).StatusCode==HttpStatusCode.Redirect,"Impresión respeta permisos de facturación pendiente");
var dashboardCalls=maintenanceFactory.Reporting.DashboardCalls;
repository.Account=repository.Account with {Permissions=["CLIENTES.CONSULTAR"]};
Check(!(await maintenanceClient.GetStringAsync("/")).Contains("hyt-dashboard-data")&&maintenanceFactory.Reporting.DashboardCalls==dashboardCalls,"Perfil sin reportes no consulta ni recibe indicadores");
using(var healthClient=factory.CreateClient(new WebApplicationFactoryClientOptions{BaseAddress=new Uri("https://localhost"),AllowAutoRedirect=false}))
{
    var health=await healthClient.GetAsync("/health/live");
    Check(health.StatusCode==HttpStatusCode.OK&&await health.Content.ReadAsStringAsync()=="ok","Sonda de vida pública confirma proceso sin exponer SQL ni secretos");
}
var proxyBuilder=Microsoft.AspNetCore.Builder.WebApplication.CreateBuilder(new Microsoft.AspNetCore.Builder.WebApplicationOptions{EnvironmentName="Production"});
proxyBuilder.Configuration["ReverseProxy:KnownProxy"]="";
var missingProxyRejected=false;
try{GestionServicios.Web.Security.ReverseProxyConfiguration.Configure(proxyBuilder);}catch(InvalidOperationException){missingProxyRejected=true;}
Check(missingProxyRejected,"Publicación rechaza proxy sin IP verificada");
proxyBuilder.Configuration["ReverseProxy:KnownProxy"]="172.30.0.1";
GestionServicios.Web.Security.ReverseProxyConfiguration.Configure(proxyBuilder);
using(var proxyServices=proxyBuilder.Services.BuildServiceProvider())
{
    var proxyOptions=proxyServices.GetRequiredService<Microsoft.Extensions.Options.IOptions<Microsoft.AspNetCore.Builder.ForwardedHeadersOptions>>().Value;
    Check(proxyOptions.ForwardLimit==1&&proxyOptions.RequireHeaderSymmetry&&proxyOptions.KnownIPNetworks.Count==0&&proxyOptions.KnownProxies.All(ip=>ip.Equals(System.Net.IPAddress.Parse("172.30.0.1"))||ip.Equals(System.Net.IPAddress.Parse("172.30.0.1").MapToIPv6())),"Solo proxy explícito puede proporcionar IP/protocolo de una solicitud");
}
var payrollAccess = new TestAccessRepository();
payrollAccess.Account = payrollAccess.Account with
{
    Permissions = ["PRESUPUESTOS.CONSULTAR", .. payrollPermissions]
};
using var payrollFactory = new HytFactory(payrollAccess);
using var payrollClient = payrollFactory.CreateClient(new WebApplicationFactoryClientOptions
    { BaseAddress = new Uri("https://localhost"), AllowAutoRedirect = false });
async Task<HttpResponseMessage> PayrollPost(string path, Dictionary<string, string> fields, string? formPath = null)
{
    var html = await payrollClient.GetStringAsync(formPath ?? path);
    var token = Regex.Match(html, "name=\"__RequestVerificationToken\"[^>]*value=\"([^\"]+)\"");
    Check(token.Success, "Formulario con protección antifalsificación: " + path);
    fields["__RequestVerificationToken"] = WebUtility.HtmlDecode(token.Groups[1].Value);
    return await payrollClient.PostAsync(path, new FormUrlEncodedContent(fields));
}
Check((await PayrollPost("/Acceso", new() { ["UserName"] = "operador", ["Password"] = payrollAccess.Password })).StatusCode
    == HttpStatusCode.Redirect, "Inicio de sesión para Planillas");
Check((await payrollClient.GetAsync("/modulo/pagos-proveedores")).Headers.Location?.OriginalString == "/Planillas/PagosProveedores",
    "El menú Planillas enlaza al listado de pagos");
Check((await PayrollPost("/Proveedores/Ficha", new()
{
    ["TaxId"] = "20000000001", ["Name"] = "Proveedor de prueba", ["Active"] = "true",
    ["Bank"] = "Banco", ["Account"] = "000012345"
})).StatusCode == HttpStatusCode.Redirect, "Alta de proveedor por formulario MVC");
const string costsUrl = "/Presupuestos/1/CostosProveedores";
var supplierFilters = await payrollClient.GetStringAsync("/Planillas/PagosProveedores?status=PAGADOS&search=Servicio&from=2026-11-01&to=2026-11-09");
Check(supplierFilters.Contains("name=\"status\"") && supplierFilters.Contains("value=\"Servicio\"") &&
    supplierFilters.IndexOf("name=\"from\"", StringComparison.Ordinal) < supplierFilters.IndexOf("name=\"status\"", StringComparison.Ordinal)
    && supplierFilters.IndexOf("name=\"status\"", StringComparison.Ordinal) < supplierFilters.IndexOf("name=\"search\"", StringComparison.Ordinal),
    "Pagos presenta fechas, estado y búsqueda en ese orden");
const string costPopupUrl = costsUrl + "/Nuevo?popup=true";
Check(!(await payrollClient.GetStringAsync(costsUrl)).Contains("name=\"SupplierId\""),
    "El listado de costos no incluye el formulario");
var budgetLookup = await payrollClient.GetStringAsync("/Planillas/PagosProveedores/ElegirPresupuesto?popup=true");
Check(budgetLookup.Contains("data-document-title=\"Elegir presupuesto\"") && budgetLookup.Contains(costsUrl + "/Nuevo"),
    "La ayuda de pagos enlaza al registro del presupuesto seleccionado");
Check((await PayrollPost(costsUrl, new()
{
    ["OperationId"] = Guid.NewGuid().ToString(), ["SupplierId"] = "1", ["Description"] = "Instalación temporal",
    ["Amount"] = "1200", ["Currency"] = "PEN", ["popup"] = "true", ["returnUrl"] = "/Planillas/PagosProveedores?all=true"
}, costPopupUrl)).StatusCode == HttpStatusCode.OK, "Registro de un costo desde el pop-up");
const string paymentUrl = "/Planillas/PagosProveedores/1";
Check((await PayrollPost(paymentUrl, new()
{
    ["OperationId"] = Guid.NewGuid().ToString(), ["Amount"] = "500", ["Date"] = "2026-11-09", ["Reference"] = "ABONO-HTTP"
})).StatusCode == HttpStatusCode.Redirect, "Abono por formulario MVC");
var paymentHtml = await payrollClient.GetStringAsync(paymentUrl);
Check(paymentHtml.Contains("ABONO-HTTP") && paymentHtml.Contains("700") && paymentHtml.Contains("000012345"),
    "El detalle muestra el abono, el saldo y la cuenta autorizada");
const string workerUrl = "/Planillas/Trabajadores/Ficha/1";
var workerHtml = await payrollClient.GetStringAsync(workerUrl);
Check(workerHtml.Contains("name=\"CargoId\"") && workerHtml.Contains("Supervisor") &&
    workerHtml.Contains("name=\"MonthlySalary\""), "La ficha ofrece selección de cargo y sueldo base");
var positionsHtml = await payrollClient.GetStringAsync("/Planillas/Cargos");
Check(!positionsHtml.Contains("Guardar cargo") && positionsHtml.Contains("/Planillas/Cargos/Ficha"),
    "El listado de cargos enlaza a la ficha y no muestra el formulario");
Check((await payrollClient.GetStringAsync("/Planillas/Cargos/Ficha")).Contains("Guardar cargo"),
    "La ficha de alta contiene el formulario");
Check((await payrollClient.GetStringAsync("/Planillas/Cargos/Ficha/1")).Contains("Supervisor"),
    "La ficha de edición carga el cargo seleccionado");
Check((await payrollClient.GetStringAsync("/Planillas/Obligaciones")).Contains("No hay obligaciones registradas"),
    "Obligaciones muestra el estado vacío sin error");
Check(workerHtml.Contains("Interbank") && workerHtml.Contains("AFP Integra"),
    "La ficha ofrece los catálogos de bancos y AFP");
const string sctrUrl = "/Planillas/SCTR";
Check((await payrollClient.GetStringAsync(sctrUrl)).Contains("Nueva póliza"), "SCTR ofrece acceso a su propia ficha");
Check((await PayrollPost(sctrUrl, new()
{
    ["Insurer"] = "Aseguradora temporal", ["Number"] = "WEB-SCTR", ["Coverage"] = "SALUD",
    ["From"] = "2026-11-01", ["To"] = "2026-11-25", ["WorkerIds[0]"] = "1", ["WorkerIds[1]"] = "2"
}, sctrUrl + "/Ficha")).StatusCode == HttpStatusCode.Redirect &&
    (await payrollClient.GetStringAsync(sctrUrl)).Contains("data-status=\"POR_VENCER\""),
    "SCTR guarda una selección múltiple y muestra su estado de vencimiento");
const string planningUrl = "/Planillas/PlanificacionPagos";
Check((await payrollClient.GetStringAsync(planningUrl + "?from=2026-11-09")).Contains("Precarga sin guardar"),
    "La planificación abre la precarga semanal");
var savePlanResponse = await PayrollPost(planningUrl, new()
{
    ["Frequency"] = "SEMANAL", ["From"] = "2026-11-09", ["Lines[0].TrabajadorId"] = "1",
    ["Lines[0].Base"] = "350", ["Lines[0].MinutosExtra"] = "60",
    ["Lines[0].ImporteExtra"] = "50", ["Lines[0].Descuento"] = "25"
});
Check(savePlanResponse.StatusCode == HttpStatusCode.Redirect &&
    (await payrollClient.GetStringAsync(planningUrl + "?from=2026-11-09")).Contains("375.00"),
    "La pantalla registra y vuelve a cargar la planificación");
Check(workerHtml.Contains("/Planillas/Trabajadores/1/Vacaciones") && !workerHtml.Contains("Registro horario"),
    "La ficha del trabajador enlaza a vacaciones y retira el acceso de registro horario");
Check((await payrollClient.GetStringAsync("/Planillas/Trabajadores")).Contains("Vacaciones pendientes"),
    "El listado muestra la columna de vacaciones pendientes");
const string vacationsUrl = "/Planillas/Trabajadores/1/Vacaciones";
var vacationHtml = await payrollClient.GetStringAsync(vacationsUrl);
Check(vacationHtml.Contains("Vacaciones pendientes") && vacationHtml.Contains("Registrar vacaciones") &&
    vacationHtml.Contains("Registrar salida de vacaciones"), "Vacaciones presenta saldo y tarjetas separadas");
Check((await PayrollPost(vacationsUrl + "/Periodos", new()
{
    ["From"] = "2025-01-01", ["To"] = "2025-12-31", ["Days"] = "10"
}, vacationsUrl)).StatusCode == HttpStatusCode.Redirect, "Registro de período de vacaciones por formulario MVC");
Check((await PayrollPost(vacationsUrl, new()
{
    ["PeriodId"] = "1", ["From"] = "2026-11-12", ["To"] = "2026-11-13", ["Days"] = "2", ["Status"] = "APROBADA"
})).StatusCode == HttpStatusCode.Redirect && (await payrollClient.GetStringAsync(vacationsUrl)).Contains("APROBADA"),
    "Las vacaciones registradas aparecen en el historial MVC");
const string attendanceUrl = "/Planillas/Trabajadores/1/Asistencia";
Check((await PayrollPost(attendanceUrl, new()
{
    ["Entry"] = "2026-11-09T08:00", ["Exit"] = "2026-11-09T17:00", ["ExtraMinutes"] = "60"
})).StatusCode == HttpStatusCode.Redirect &&
    (await payrollClient.GetStringAsync(attendanceUrl + "?from=2026-11-01&to=2026-11-30")).Contains("17:00"),
    "La jornada registrada muestra la salida y el total de extras");

sealed class HytFactory(TestAccessRepository repository,bool turnstileEnabled=false,string siteKey="sitekey-prueba") : WebApplicationFactory<WebEntryPoint>
{
    public StubCloudflareHandler Turnstile {get;}=new();
    public string FileRoot {get;}=Path.Combine(Path.GetTempPath(),"hyt-http-files-"+Guid.NewGuid().ToString("N"));
    public TestInvoiceMarkRepository InvoiceMarks {get;}=new();
    public TestReportingRepository Reporting {get;}=new();
    public TestClock Clock { get; } = new();
    public TestMaintenanceRepository Maintenance { get; } = new();
    public TestTaxRateRepository TaxRates { get; } = new();
    public TestBudgetRepository Budgets { get; } = new();
    public TestDocumentRepository Documents {get;}=new();
    public TestPrintRepository Prints {get;}=new();
    public TestCompanyRepository Company {get;}=new();
    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if(!disposing||!Directory.Exists(FileRoot))return;
        var absolute=Path.GetFullPath(FileRoot);var parent=Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
        if(Path.GetDirectoryName(absolute)!=parent||!Regex.IsMatch(Path.GetFileName(absolute),@"^hyt-http-files-[a-f0-9]{32}$"))throw new InvalidOperationException("Ruta de archivos de prueba inesperada.");
        Directory.Delete(absolute,true);
    }
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "SistemaWebHYT.slnx"))) directory = directory.Parent;
        builder.UseEnvironment("Development").UseContentRoot(Path.Combine(directory!.FullName, "GestionServicios.Web"));
        builder.ConfigureAppConfiguration((_,configuration)=>configuration.AddInMemoryCollection(new Dictionary<string,string?>{["DocumentStorage:Root"]=FileRoot,["CloudflareTurnstile:Enabled"]=turnstileEnabled.ToString(),["CloudflareTurnstile:SiteKey"]=siteKey,["CloudflareTurnstile:SecretKey"]="secret-prueba",["CloudflareTurnstile:AllowedHostnames:0"]="localhost"}));
        builder.ConfigureLogging(logging => logging.SetMinimumLevel(LogLevel.Error));
        builder.ConfigureServices(services =>
        {
            services.AddHttpClient<GestionServicios.Web.Security.ITurnstileVerifier,GestionServicios.Web.Security.CloudflareTurnstileVerifier>().ConfigurePrimaryHttpMessageHandler(()=>Turnstile);
            services.AddControllersWithViews().AddApplicationPart(typeof(UiFixtureController).Assembly);
            services.RemoveAll<IAccessRepository>(); services.AddSingleton<IAccessRepository>(repository);
            services.RemoveAll<ISecurityAdministrationRepository>();
            services.AddSingleton<ISecurityAdministrationRepository>(new TestSecurityAdministration(repository));
            services.RemoveAll<IMaintenanceRepository>(); services.AddSingleton<IMaintenanceRepository>(Maintenance);
            services.RemoveAll<ITaxRateRepository>(); services.AddSingleton<ITaxRateRepository>(TaxRates);
            services.RemoveAll<GestionServicios.Web.Services.IDocumentFileStore>();services.AddScoped<GestionServicios.Web.Services.IDocumentFileStore,GestionServicios.Web.Services.PrivateDocumentFileStore>();
            services.RemoveAll<GestionServicios.Application.Followup.IDocumentFileRepository>();services.AddSingleton<GestionServicios.Application.Followup.IDocumentFileRepository>(new TestDocumentFileRepository(Documents));
            services.RemoveAll<GestionServicios.Application.Followup.IDocumentRepository>(); services.AddSingleton<GestionServicios.Application.Followup.IDocumentRepository>(Documents);
            services.RemoveAll<GestionServicios.Application.Billing.IInvoiceMarkRepository>();services.AddSingleton<GestionServicios.Application.Billing.IInvoiceMarkRepository>(InvoiceMarks);
            services.RemoveAll<GestionServicios.Application.Reporting.IReportingRepository>();services.AddSingleton<GestionServicios.Application.Reporting.IReportingRepository>(Reporting);
            services.RemoveAll<IBudgetRepository>(); services.AddSingleton<IBudgetRepository>(Budgets);
            services.RemoveAll<IBudgetPrintRepository>();services.AddSingleton<IBudgetPrintRepository>(Prints);
            services.RemoveAll<ICompanySettingsRepository>();services.AddSingleton<ICompanySettingsRepository>(Company);
            services.RemoveAll<TimeProvider>(); services.AddSingleton<TimeProvider>(Clock);
            var payroll = new TestPayrollRepository();
            services.RemoveAll<ISupplierPaymentsRepository>(); services.AddSingleton<ISupplierPaymentsRepository>(payroll);
            services.RemoveAll<IWorkersRepository>(); services.AddSingleton<IWorkersRepository>(payroll);
            services.RemoveAll<IPersonalPaymentsRepository>(); services.AddSingleton<IPersonalPaymentsRepository>(payroll);
            services.RemoveAll<IPayrollObligationsRepository>(); services.AddSingleton<IPayrollObligationsRepository>(payroll);
            services.RemoveAll<IPayPlanningRepository>(); services.AddSingleton<IPayPlanningRepository>(payroll);
            services.RemoveAll<IPayrollCatalogRepository>(); services.AddSingleton<IPayrollCatalogRepository>(payroll);
        });
    }
}

sealed class TestAccessRepository : IAccessRepository
{
    public string Password { get; } = Guid.NewGuid().ToString("N") + "aA1!";
    private AccessAccount account = null!;
    public AccessAccount Account
    {
        get => account;
        set
        {
            var permissions = value.Permissions.ToHashSet(StringComparer.Ordinal);
            foreach (var module in new[] { "CLIENTES", "TIENDAS", "MATERIALES", "UNIDADES", "ZONAS", "TARIFARIOS",
                "PRESUPUESTOS", "OC_CLIENTE", "GUIAS", "CONFIRMACIONES", "EJECUCION" })
            {
                if (permissions.Contains(module + ".EDITAR"))
                {
                    permissions.Add(module + ".CREAR");
                    permissions.Add(module + ".MODIFICAR");
                    permissions.Add(module + ".ACTIVAR");
                    permissions.Add(module + ".CAMBIAR_ESTADO");
                }
                if (permissions.Contains(module + ".CONSULTAR")) permissions.Add(module + ".DESCARGAR");
            }
            if (permissions.Contains("REPORTES.CONSULTAR"))
            {
                permissions.Add("REPORTES.IMPRIMIR");
                if (permissions.Contains("PRESUPUESTOS.CONSULTAR")) permissions.Add("PRESUPUESTOS.IMPRIMIR");
            }
            account = value with { Permissions = permissions.ToArray() };
        }
    }
    public int NameLookups {get;private set;}
    public int Failures { get; private set; }
    public int Logouts { get; private set; }
    public bool Unavailable { get; set; }
    public TestAccessRepository()
    {
        Account = new("test-account", "operador", "Operador HYT", null, Guid.NewGuid().ToString("N"), true, null, ["Consulta"], ["CLIENTES.CONSULTAR"]);
        Account = Account with { PasswordHash = new PasswordHasher<AccessAccount>().HashPassword(Account, Password) };
    }
    public int LookupDelay {get;set;}
    public async Task<AccessAccount?> FindByNameAsync(string name, CancellationToken cancellationToken = default) {NameLookups++;if(LookupDelay>0)await Task.Delay(LookupDelay,cancellationToken);return name == "OPERADOR" ? Account : null;}
    public Task<AccessAccount?> FindByIdAsync(string id, CancellationToken cancellationToken = default)
        => Unavailable ? throw new InvalidOperationException("Repositorio de pruebas no disponible") : Task.FromResult<AccessAccount?>(id == Account.Id ? Account : null);
    public Task<bool> CompleteAttemptAsync(string id, string stamp, bool valid, string? upgradedHash, Guid correlationId, CancellationToken cancellationToken = default)
    {
        if (!valid) Failures++;
        return Task.FromResult(valid && Account.Active && stamp == Account.SecurityStamp && (Account.LockoutEnd is null || Account.LockoutEnd <= DateTimeOffset.UtcNow));
    }
    public Task RecordLogoutAsync(string id, Guid correlationId, CancellationToken cancellationToken = default) { Logouts++; return Task.CompletedTask; }
}

sealed class TestClock : TimeProvider
{
    private TimeSpan offset;
    public override DateTimeOffset GetUtcNow() => DateTimeOffset.UtcNow + offset;
    public void Advance(TimeSpan value) => offset += value;
    public void Reset() => offset = TimeSpan.Zero;
}
