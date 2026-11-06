// Firma: FRANCO LARA - 05/11/2026 | Registra captcha manual acotado y separa límite de imágenes/renovación de los intentos de acceso.
using System.Globalization;
using Microsoft.AspNetCore.Localization;
using System.Threading.RateLimiting;
using GestionServicios.Application.Security;
using GestionServicios.Infrastructure;
using GestionServicios.Infrastructure.Security;
using GestionServicios.Web.Security;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Identity;

var builder = WebApplication.CreateBuilder(args);
ReverseProxyConfiguration.Configure(builder);
builder.Services.AddControllersWithViews(options => options.Filters.Add(new Microsoft.AspNetCore.Mvc.AutoValidateAntiforgeryTokenAttribute()));
builder.Services.Configure<Microsoft.AspNetCore.Mvc.MvcOptions>(options => options.Filters.Add<GestionServicios.Web.MaintenanceUnavailableFilter>());
builder.Services.Configure<RequestLocalizationOptions>(options =>
{
    var culture = new CultureInfo("es-PE");
    options.DefaultRequestCulture = new RequestCulture(culture);
    options.SupportedCultures = [culture];
    options.SupportedUICultures = [culture];
});
builder.Services.AddSingleton<TimeProvider>(TimeProvider.System);
builder.Services.AddScoped(_ => new SqlConnectionFactory(builder.Configuration.GetConnectionString("HYT") ?? ""));
builder.Services.AddScoped<IAccessRepository, SqlAccessRepository>();
builder.Services.AddScoped<GestionServicios.Application.Maintenance.IMaintenanceRepository, GestionServicios.Infrastructure.Maintenance.SqlMaintenanceRepository>();
builder.Services.AddScoped<GestionServicios.Application.Maintenance.ITaxRateRepository, GestionServicios.Infrastructure.Maintenance.SqlTaxRateRepository>();
builder.Services.AddScoped<GestionServicios.Application.Maintenance.ICompanySettingsRepository, GestionServicios.Infrastructure.Maintenance.SqlCompanySettingsRepository>();
builder.Services.AddScoped<GestionServicios.Application.Followup.IDocumentFileRepository, GestionServicios.Infrastructure.Followup.SqlDocumentFileRepository>();
builder.Services.Configure<GestionServicios.Web.Services.CloudflareR2Settings>(builder.Configuration.GetSection("CloudflareR2"));
builder.Services.AddScoped<GestionServicios.Web.Services.IR2ObjectClient,GestionServicios.Web.Services.CloudflareR2ObjectClient>();
builder.Services.AddScoped<GestionServicios.Web.Services.IDocumentFileStore,GestionServicios.Web.Services.R2DocumentFileStore>();
builder.Services.AddScoped<GestionServicios.Application.Reporting.IReportingRepository,GestionServicios.Infrastructure.Reporting.SqlReportingRepository>();
builder.Services.AddScoped<GestionServicios.Application.Billing.IInvoiceMarkRepository,GestionServicios.Infrastructure.Billing.SqlInvoiceMarkRepository>();
builder.Services.AddScoped<GestionServicios.Application.Followup.IDocumentRepository, GestionServicios.Infrastructure.Followup.SqlDocumentRepository>();
builder.Services.AddScoped<GestionServicios.Application.Budgets.IBudgetRepository, GestionServicios.Infrastructure.Budgets.SqlBudgetRepository>();
builder.Services.AddScoped<GestionServicios.Application.Budgets.IBudgetPrintRepository, GestionServicios.Infrastructure.Budgets.SqlBudgetPrintRepository>();
builder.Services.AddScoped<GestionServicios.Web.Services.BudgetPrintRenderer>();
builder.Services.AddScoped<IPasswordHasher<AccessAccount>, PasswordHasher<AccessAccount>>();
builder.Services.Configure<CloudflareTurnstileSettings>(builder.Configuration.GetSection("CloudflareTurnstile"));
builder.Services.AddHttpClient<ITurnstileVerifier,CloudflareTurnstileVerifier>(client=>client.Timeout=TimeSpan.FromSeconds(10));
builder.Services.AddSingleton<ManualCaptchaService>();
builder.Services.AddScoped<AccessService>();
builder.Services.AddScoped<HytCookieEvents>();
builder.Services.AddAuthentication("HytCookie").AddCookie("HytCookie", options =>
{
    options.Cookie.Name = "__Host-Hyt.Session";
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.HttpOnly = true;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.LoginPath = "/Acceso";
    options.AccessDeniedPath = "/Denegado";
    options.ExpireTimeSpan = TimeSpan.FromMinutes(30);
    options.SlidingExpiration = true;
    options.EventsType = typeof(HytCookieEvents);
});
builder.Services.AddAuthorization(options =>
{
    options.FallbackPolicy = new AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build();
    foreach (var permission in new[] { "CLIENTES.CONSULTAR", "USUARIOS.CONSULTAR", "CONFIGURACION.CONSULTAR", "AUDITORIA.CONSULTAR" })
        options.AddPolicy(permission, policy => policy.RequireClaim(AccessService.PermissionClaim, permission));
});
builder.Services.AddSingleton<IAuthorizationPolicyProvider, PermissionPolicyProvider>();
builder.Services.AddRateLimiter(options =>
{
    options.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    options.OnRejected = async (context, cancellationToken) =>
    {
        context.HttpContext.Response.Headers.RetryAfter = "60";
        context.HttpContext.Response.ContentType = "text/plain; charset=utf-8";
        await context.HttpContext.Response.WriteAsync("Demasiados intentos de acceso. Espera un minuto e inténtalo nuevamente.", cancellationToken);
    };
    options.AddPolicy("captcha", context => RateLimitPartition.GetFixedWindowLimiter(
        context.Connection.RemoteIpAddress?.ToString() ?? "local", _ => new FixedWindowRateLimiterOptions
        { PermitLimit = 120, Window = TimeSpan.FromMinutes(1), QueueLimit = 0 }));
    options.AddPolicy("access", context => RateLimitPartition.GetFixedWindowLimiter(
        context.Connection.RemoteIpAddress?.ToString() ?? "local", _ => new FixedWindowRateLimiterOptions
        { PermitLimit = 20, Window = TimeSpan.FromMinutes(1), QueueLimit = 0 }));
});
var protection = builder.Services.AddDataProtection().SetApplicationName("SistemaWebHYT");
var keysPath = builder.Configuration["Security:DataProtectionPath"];
if (!string.IsNullOrWhiteSpace(keysPath)) protection.PersistKeysToFileSystem(new DirectoryInfo(keysPath));
else if (!builder.Environment.IsDevelopment())
    throw new InvalidOperationException("Configure Security:DataProtectionPath para conservar las claves de sesión.");

var app = builder.Build();
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    app.UseHsts();
}
app.UseForwardedHeaders();
app.UseHttpsRedirection();
app.Use(async (context, next) =>
{
    context.Response.OnStarting(() =>
    {
        if (context.Response.ContentType?.StartsWith("text/html", StringComparison.OrdinalIgnoreCase) == true)
            context.Response.Headers.CacheControl = "no-store";
        context.Response.Headers["X-Content-Type-Options"] = "nosniff";
        return Task.CompletedTask;
    });
    await next();
});
app.UseRequestLocalization();
app.UseAuthentication();
app.UseAuthorization();
app.UseRateLimiter();
app.UseAntiforgery();
app.MapGet("/health/live",()=>Results.Text("ok","text/plain")).AllowAnonymous();
app.MapStaticAssets().AllowAnonymous();
app.MapControllerRoute(name: "default", pattern: "{controller=Home}/{action=Index}/{id?}");
app.Run();

public partial class Program;
