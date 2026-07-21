using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Localization;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.AspNetCore.CookiePolicy;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.EntityFrameworkCore;
using SistemaControlEspaciosDeportivosWeb.Data;
using SistemaControlEspaciosDeportivosWeb.Configuration;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using System.Net.Http.Headers;
using System.Globalization;
using System.IO;
using System.Threading.RateLimiting;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<BusinessInformationOptions>(builder.Configuration.GetSection(BusinessInformationOptions.SectionName));
builder.Services.Configure<LegalDocumentsOptions>(builder.Configuration.GetSection(LegalDocumentsOptions.SectionName));

// En desarrollo, forzamos User Secrets al final para que tenga prioridad
// sobre valores anteriores de appsettings/perfiles locales.
if (builder.Environment.IsDevelopment())
{
    builder.Configuration.AddUserSecrets<Program>(optional: true, reloadOnChange: true);
}

var dataProtectionKeysPath = (builder.Configuration["DataProtection:KeysPath"] ?? string.Empty).Trim();
if (string.IsNullOrWhiteSpace(dataProtectionKeysPath))
{
    dataProtectionKeysPath = Path.Combine(builder.Environment.ContentRootPath, "App_Data", "DataProtection-Keys");
}
Directory.CreateDirectory(dataProtectionKeysPath);
builder.Services.AddDataProtection()
    .SetApplicationName("SistemaControlEspaciosDeportivosWeb")
    .PersistKeysToFileSystem(new DirectoryInfo(dataProtectionKeysPath));

// Add services to the container.
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(connectionString));
builder.Services.AddDatabaseDeveloperPageExceptionFilter();

var identityBehaviorSettings = builder.Configuration
    .GetSection(IdentityBehaviorSettings.SectionName)
    .Get<IdentityBehaviorSettings>() ?? new IdentityBehaviorSettings();
builder.Services.Configure<IdentityBehaviorSettings>(
    builder.Configuration.GetSection(IdentityBehaviorSettings.SectionName));

builder.Services.AddDefaultIdentity<ApplicationUser>(options =>
    {
        options.SignIn.RequireConfirmedAccount = identityBehaviorSettings.RequireConfirmedAccount;

        options.Password.RequiredLength = 8;
        options.Password.RequireDigit = true;
        options.Password.RequireLowercase = true;
        options.Password.RequireUppercase = true;
        options.Password.RequireNonAlphanumeric = true;
        options.Password.RequiredUniqueChars = 1;

        options.Lockout.AllowedForNewUsers = true;
        options.Lockout.MaxFailedAccessAttempts = 5;
        options.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(15);
    })
    .AddErrorDescriber<SpanishIdentityErrorDescriber>()
    .AddRoles<IdentityRole>()
    .AddEntityFrameworkStores<ApplicationDbContext>();

var googleClientId = (builder.Configuration["Authentication:Google:ClientId"] ?? string.Empty).Trim();
var googleClientSecret = (builder.Configuration["Authentication:Google:ClientSecret"] ?? string.Empty).Trim();
if (!string.IsNullOrWhiteSpace(googleClientId) && !string.IsNullOrWhiteSpace(googleClientSecret))
{
    builder.Services.AddAuthentication()
        .AddGoogle("Google", options =>
        {
            options.ClientId = googleClientId;
            options.ClientSecret = googleClientSecret;
            options.SignInScheme = IdentityConstants.ExternalScheme;
        });
}
builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Identity/Account/Login";
    options.AccessDeniedPath = "/Identity/Account/AccessDenied";
    options.ExpireTimeSpan = TimeSpan.FromMinutes(30);
    options.SlidingExpiration = true;
    options.Cookie.HttpOnly = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.Cookie.IsEssential = true;
});
builder.Services.Configure<CookiePolicyOptions>(options =>
{
    options.MinimumSameSitePolicy = SameSiteMode.Lax;
    options.HttpOnly = HttpOnlyPolicy.Always;
    options.Secure = CookieSecurePolicy.Always;
});
builder.Services.AddHsts(options =>
{
    options.Preload = true;
    options.IncludeSubDomains = true;
    options.MaxAge = TimeSpan.FromDays(365);
});
builder.Services.AddHttpsRedirection(options =>
{
    options.RedirectStatusCode = StatusCodes.Status308PermanentRedirect;
});
builder.Services.AddLocalization(options => options.ResourcesPath = "Resources");
builder.Services.AddControllersWithViews(options =>
    {
        options.Filters.AddService<OnboardingGuardFilter>();
        options.ModelBindingMessageProvider.SetValueMustNotBeNullAccessor(_ => "Este campo es obligatorio.");
        options.ModelBindingMessageProvider.SetAttemptedValueIsInvalidAccessor((valor, campo) => $"El valor '{valor}' no es valido para {campo}.");
        options.ModelBindingMessageProvider.SetMissingBindRequiredValueAccessor(campo => $"El campo {campo} es obligatorio.");
        options.ModelBindingMessageProvider.SetMissingKeyOrValueAccessor(() => "Se requiere un valor.");
        options.ModelBindingMessageProvider.SetMissingRequestBodyRequiredValueAccessor(() => "El cuerpo de la solicitud es obligatorio.");
        options.ModelBindingMessageProvider.SetNonPropertyAttemptedValueIsInvalidAccessor(valor => $"El valor '{valor}' no es valido.");
        options.ModelBindingMessageProvider.SetNonPropertyUnknownValueIsInvalidAccessor(() => "El valor proporcionado no es valido.");
        options.ModelBindingMessageProvider.SetNonPropertyValueMustBeANumberAccessor(() => "El valor debe ser numerico.");
        options.ModelBindingMessageProvider.SetUnknownValueIsInvalidAccessor(campo => $"El valor proporcionado no es valido para {campo}.");
        options.ModelBindingMessageProvider.SetValueIsInvalidAccessor(valor => $"El valor '{valor}' no es valido.");
        options.ModelBindingMessageProvider.SetValueMustBeANumberAccessor(campo => $"El campo {campo} debe ser numerico.");
    })
    .AddViewLocalization()
    .AddDataAnnotationsLocalization();
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.Cookie.IsEssential = true;
    options.IdleTimeout = TimeSpan.FromMinutes(20);
});
builder.Services.AddAntiforgery(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
    options.Cookie.SameSite = SameSiteMode.Strict;
});
builder.Services.AddRateLimiter(options =>
{
    options.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    options.OnRejected = async (context, _) =>
    {
        context.HttpContext.Response.ContentType = "application/json; charset=utf-8";
        await context.HttpContext.Response.WriteAsync("{\"ok\":false,\"mensaje\":\"Demasiadas solicitudes. Intenta nuevamente en unos segundos.\"}");
    };

    options.GlobalLimiter = PartitionedRateLimiter.Create<HttpContext, string>(httpContext =>
    {
        static string KeyByIp(HttpContext context, string scope)
        {
            var ip = context.Connection.RemoteIpAddress?.ToString();
            if (string.IsNullOrWhiteSpace(ip))
                ip = "sin-ip";
            return $"{scope}:{ip}";
        }

        var path = httpContext.Request.Path.Value?.ToLowerInvariant() ?? string.Empty;
        var isPost = HttpMethods.IsPost(httpContext.Request.Method);

        if (isPost && path is "/identity/account/login")
        {
            return RateLimitPartition.GetFixedWindowLimiter(
                partitionKey: KeyByIp(httpContext, "login"),
                factory: _ => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = 5,
                    Window = TimeSpan.FromMinutes(1),
                    QueueProcessingOrder = QueueProcessingOrder.OldestFirst,
                    QueueLimit = 0,
                    AutoReplenishment = true
                });
        }

        if (isPost && path is "/identity/account/register")
        {
            return RateLimitPartition.GetFixedWindowLimiter(
                partitionKey: KeyByIp(httpContext, "register"),
                factory: _ => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = 3,
                    Window = TimeSpan.FromMinutes(10),
                    QueueProcessingOrder = QueueProcessingOrder.OldestFirst,
                    QueueLimit = 0,
                    AutoReplenishment = true
                });
        }

        if (isPost && path is "/identity/account/forgotpassword")
        {
            return RateLimitPartition.GetFixedWindowLimiter(
                partitionKey: KeyByIp(httpContext, "forgot"),
                factory: _ => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = 3,
                    Window = TimeSpan.FromMinutes(15),
                    QueueProcessingOrder = QueueProcessingOrder.OldestFirst,
                    QueueLimit = 0,
                    AutoReplenishment = true
                });
        }

        if (isPost && path is "/home/crearreservapublica" or "/home/solicitarreservapublica")
        {
            return RateLimitPartition.GetFixedWindowLimiter(
                partitionKey: KeyByIp(httpContext, "reserva-publica"),
                factory: _ => new FixedWindowRateLimiterOptions
                {
                    PermitLimit = 6,
                    Window = TimeSpan.FromMinutes(1),
                    QueueProcessingOrder = QueueProcessingOrder.OldestFirst,
                    QueueLimit = 0,
                    AutoReplenishment = true
                });
        }

        return RateLimitPartition.GetNoLimiter("sin-limite");
    });
});
builder.Services.AddHttpContextAccessor();
builder.Services.Configure<BrevoSettings>(builder.Configuration.GetSection("Brevo"));
builder.Services.Configure<AutomationSettings>(builder.Configuration.GetSection("AutomationSettings"));
builder.Services.Configure<JobsSettings>(builder.Configuration.GetSection("Jobs"));
builder.Services.Configure<SedeImagenStorageSettings>(builder.Configuration.GetSection("SedeImagenStorage"));
builder.Services.Configure<CloudflareTurnstileSettings>(builder.Configuration.GetSection("CloudflareTurnstile"));
builder.Services.AddHttpClient<IEmailService, BrevoEmailService>(httpClient =>
{
    httpClient.BaseAddress = new Uri("https://api.brevo.com/v3/");
    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
});
builder.Services.AddHttpClient("GooglePlacesTextSearch", httpClient =>
{
    httpClient.BaseAddress = new Uri("https://maps.googleapis.com/");
    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
});
builder.Services.AddHttpClient();
builder.Services.AddHttpClient<ITurnstileValidationService, TurnstileValidationService>();
builder.Services.AddScoped<IModuloPermisoService, ModuloPermisoService>();
builder.Services.AddScoped<ISportCenterStoredProcedureService, SportCenterStoredProcedureService>();
builder.Services.AddScoped<OnboardingGuardFilter>();
builder.Services.AddScoped<IComprobanteElectronicoEmisionService, ComprobanteElectronicoEmisionService>();
builder.Services.AddScoped<IHomeReferencialesExternosSyncService, HomeReferencialesExternosSyncService>();
builder.Services.AddScoped<ISedeImagenStorageService, R2SedeImagenStorageService>();
builder.Services.AddScoped<IAccountEmailService, AccountEmailService>();
builder.Services.AddScoped<IClubRegistrationNotificationService, ClubRegistrationNotificationService>();
builder.Services.AddScoped<IReservationEmailNotificationService, ReservationEmailNotificationService>();
builder.Services.AddScoped<IDesafioEmailNotificationService, DesafioEmailNotificationService>();
builder.Services.AddHostedService<ReservaAutomationHostedService>();

var app = builder.Build();

var supportedCultures = new[]
{
    new CultureInfo("es-PE"),
    new CultureInfo("es-419")
};

var localizationOptions = new RequestLocalizationOptions
{
    DefaultRequestCulture = new RequestCulture("es-PE"),
    SupportedCultures = supportedCultures,
    SupportedUICultures = supportedCultures
};

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseMigrationsEndPoint();
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();

    app.MapWhen(context => context.Request.Path.StartsWithSegments("/dev"), branch =>
    {
        branch.Run(async context =>
        {
            context.Response.StatusCode = StatusCodes.Status404NotFound;
            await context.Response.WriteAsync("Not Found");
        });
    });
}

app.UseHttpsRedirection();
app.Use(async (context, next) =>
{
    var path = context.Request.Path.Value ?? string.Empty;
    if (path.Equals("/Home", StringComparison.OrdinalIgnoreCase)
        || path.Equals("/Home/Index", StringComparison.OrdinalIgnoreCase))
    {
        var destino = "/" + context.Request.QueryString;
        context.Response.Redirect(destino, permanent: true);
        return;
    }

    await next();
});
app.UseRequestLocalization(localizationOptions);
app.UseRouting();
app.UseCookiePolicy();
app.UseRateLimiter();
app.UseSession();
app.UseStaticFiles();

app.UseAuthentication();
app.UseAuthorization();
app.UseStatusCodePages(async context =>
{
    var http = context.HttpContext;
    if (http.Response.StatusCode == StatusCodes.Status400BadRequest
        && http.Request.Path.StartsWithSegments("/Identity/Account/Logout", StringComparison.OrdinalIgnoreCase))
    {
        http.Response.Redirect("/Identity/Account/Login?sessionExpired=1");
        return;
    }

    await Task.CompletedTask;
});

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

await IdentitySeeder.SeedRolesAsync(app.Services);

app.Run();
