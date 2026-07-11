using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.EventLog;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Data;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = ResolverContentRoot()
});
if (OperatingSystem.IsWindows())
{
    builder.Logging.AddFilter<EventLogLoggerProvider>(level => level >= LogLevel.None);
}

var secretsRootPath = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
    "Microsoft",
    "UserSecrets",
    "aspnet-SistemaAdministrativoWeb-53c45fe8-ede4-4ae5-9d55-e176dba84e4e");
var secretsPath = Path.Combine(secretsRootPath, "secrets.json");
var secretsLocalPath = Path.Combine(secretsRootPath, "secretsLocal.json");

if (builder.Environment.IsDevelopment())
{
    builder.Configuration.AddJsonFile(secretsPath, optional: true, reloadOnChange: true);
    builder.Configuration.AddJsonFile(secretsLocalPath, optional: true, reloadOnChange: true);
    builder.Configuration.AddUserSecrets<Program>(optional: true, reloadOnChange: true);
}

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection")
    ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(connectionString));
builder.Services.AddDatabaseDeveloperPageExceptionFilter();
builder.Services.Configure<IdentitySeedOptions>(
    builder.Configuration.GetSection(IdentitySeedOptions.SectionName));
builder.Services.Configure<CloudflareTurnstileSettings>(
    builder.Configuration.GetSection(CloudflareTurnstileSettings.SectionName));
builder.Services.Configure<MigoApiSettings>(
    builder.Configuration.GetSection(MigoApiSettings.SectionName));
builder.Services.Configure<BusinessInformationOptions>(
    builder.Configuration.GetSection(BusinessInformationOptions.SectionName));
builder.Services.Configure<LegalDocumentsOptions>(
    builder.Configuration.GetSection(LegalDocumentsOptions.SectionName));
var identityBehaviorSettings = builder.Configuration
    .GetSection(IdentityBehaviorSettings.SectionName)
    .Get<IdentityBehaviorSettings>() ?? new IdentityBehaviorSettings();
builder.Services.Configure<IdentityBehaviorSettings>(
    builder.Configuration.GetSection(IdentityBehaviorSettings.SectionName));

builder.Services.AddDefaultIdentity<IdentityUser>(options =>
    {
        options.SignIn.RequireConfirmedAccount = identityBehaviorSettings.RequireConfirmedAccount;
        options.Password.RequiredLength = 6;
        options.Password.RequireDigit = true;
        options.Password.RequireLowercase = true;
        options.Password.RequireUppercase = true;
        options.Password.RequireNonAlphanumeric = true;
    })
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

builder.Services.AddControllersWithViews();
builder.Services.AddMemoryCache();
builder.Services.AddHttpContextAccessor();
builder.Services.AddSession(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
    options.IdleTimeout = TimeSpan.FromHours(8);
});

builder.Services.AddScoped<IDbConnectionFactory, SqlConnectionFactory>();
builder.Services.AddScoped<IPlanCuentaRepository, PlanCuentaRepository>();
builder.Services.AddScoped<IDiferenciaCambioRepository, DiferenciaCambioRepository>();
builder.Services.AddScoped<IAjusteCuentaRepository, AjusteCuentaRepository>();
builder.Services.AddScoped<IAperturaProcesoRepository, AperturaProcesoRepository>();
builder.Services.AddScoped<ICierreProcesoRepository, CierreProcesoRepository>();
builder.Services.AddScoped<IBancoRepository, BancoRepository>();
builder.Services.AddScoped<ICentroCostoRepository, CentroCostoRepository>();
builder.Services.AddScoped<ICuentaCorrienteRepository, CuentaCorrienteRepository>();
builder.Services.AddScoped<ICajaBancoRepository, CajaBancoRepository>();
builder.Services.AddScoped<IOrigenRepository, OrigenRepository>();
builder.Services.AddScoped<ICuentaDestinoReglaRepository, CuentaDestinoReglaRepository>();
builder.Services.AddScoped<IConfiguracionContabilizacionRepository, ConfiguracionContabilizacionRepository>();
builder.Services.AddScoped<IAsientoPreviewService, AsientoPreviewService>();
builder.Services.AddScoped<IMonedaRepository, MonedaRepository>();
builder.Services.AddScoped<ITipoCambioRepository, TipoCambioRepository>();
builder.Services.AddScoped<IPeriodoContableRepository, PeriodoContableRepository>();
builder.Services.AddScoped<IPeriodoContableService, PeriodoContableService>();
builder.Services.AddScoped<ITipoCambioSyncService, TipoCambioSyncService>();
builder.Services.AddScoped<IAnalisisCuentaRepository, AnalisisCuentaRepository>();
builder.Services.AddScoped<IBalanceComprobacionRepository, BalanceComprobacionRepository>();
builder.Services.AddScoped<IRegistroVentasRepository, RegistroVentasRepository>();
builder.Services.AddScoped<IRegistroComprasRepository, RegistroComprasRepository>();
builder.Services.AddScoped<ILibroDiarioRepository, LibroDiarioRepository>();
builder.Services.AddScoped<ILibroMayorRepository, LibroMayorRepository>();
builder.Services.AddScoped<ILibroElectronicoRepository, LibroElectronicoRepository>();
builder.Services.AddScoped<ILibroDiario51Service, LibroDiario51Service>();
builder.Services.AddScoped<ILibroDiario52Service, LibroDiario52Service>();
builder.Services.AddScoped<ILibroMayor61Service, LibroMayor61Service>();
builder.Services.AddScoped<IPleFileNameService, PleFileNameService>();
builder.Services.AddScoped<IPleValidationService, PleValidationService>();
builder.Services.AddScoped<IPleTxtGenerator, PleTxtGenerator>();
builder.Services.AddScoped<IPleDownloadStore, PleDownloadStore>();
builder.Services.AddScoped<ILibroElectronicoService, LibroElectronicoService>();
builder.Services.AddHttpClient<IMigoTipoCambioApiClient, MigoTipoCambioApiClient>((serviceProvider, httpClient) =>
{
    var settings = serviceProvider.GetRequiredService<Microsoft.Extensions.Options.IOptions<MigoApiSettings>>().Value;
    if (!string.IsNullOrWhiteSpace(settings.BaseUrl))
    {
        httpClient.BaseAddress = new Uri(settings.BaseUrl, UriKind.Absolute);
    }
});
builder.Services.AddHttpClient<IMigoPadronApiClient, MigoPadronApiClient>((serviceProvider, httpClient) =>
{
    var settings = serviceProvider.GetRequiredService<Microsoft.Extensions.Options.IOptions<MigoApiSettings>>().Value;
    if (!string.IsNullOrWhiteSpace(settings.BaseUrl))
    {
        httpClient.BaseAddress = new Uri(settings.BaseUrl, UriKind.Absolute);
    }
});
builder.Services.AddScoped<IAsientoRepository, AsientoRepository>();
builder.Services.AddScoped<IDetraccionSunatRepository, DetraccionSunatRepository>();
builder.Services.AddScoped<ITipoPercepcionRepository, TipoPercepcionRepository>();
builder.Services.AddScoped<IProveedorRepository, ProveedorRepository>();
builder.Services.AddScoped<ICompraRepository, CompraRepository>();
builder.Services.AddScoped<IXmlProvisionImportService, XmlProvisionImportService>();
builder.Services.AddScoped<IAplicacionNotaCreditoRepository, AplicacionNotaCreditoRepository>();
builder.Services.AddScoped<ITipoComprobanteRepository, TipoComprobanteRepository>();
builder.Services.AddScoped<ITipoAfectacionIgvRepository, TipoAfectacionIgvRepository>();
builder.Services.AddScoped<IClienteRepository, ClienteRepository>();
builder.Services.AddScoped<IPersonaRepository, PersonaRepository>();
builder.Services.AddScoped<IVentaRepository, VentaRepository>();
builder.Services.AddScoped<IEmpresaRepository, EmpresaRepository>();
builder.Services.AddScoped<ICurrentCompanyAccessor, SessionCurrentCompanyAccessor>();
builder.Services.AddScoped<ICuentaAdministradoraRepository, CuentaAdministradoraRepository>();
builder.Services.AddScoped<IModulePermissionService, ModulePermissionService>();
builder.Services.AddScoped<IParametroEmpresaRepository, ParametroEmpresaRepository>();
builder.Services.AddScoped<IdentityStartupSeeder>();
builder.Services.AddHttpClient<ITurnstileValidationService, TurnstileValidationService>();

var app = builder.Build();

await using (var scope = app.Services.CreateAsyncScope())
{
    var seeder = scope.ServiceProvider.GetRequiredService<IdentityStartupSeeder>();
    await seeder.SeedAsync();
}

if (app.Environment.IsDevelopment())
{
    app.UseMigrationsEndPoint();
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRouting();
app.UseSession();
app.UseAuthentication();
app.UseMiddleware<ActiveCompanySessionValidationMiddleware>();
app.UseAuthorization();

app.MapStaticAssets();

app.MapControllerRoute(
        name: "areas",
        pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}")
    .WithStaticAssets();

app.MapControllerRoute(
        name: "default",
        pattern: "{controller=Home}/{action=Index}/{id?}")
    .WithStaticAssets();

app.MapRazorPages()
    .WithStaticAssets();

app.Run();

static string ResolverContentRoot()
{
    var directorioActual = Directory.GetCurrentDirectory();
    if (ExisteEstructuraProyecto(directorioActual))
    {
        return directorioActual;
    }

    var candidato = new DirectoryInfo(AppContext.BaseDirectory);
    while (candidato is not null)
    {
        if (ExisteEstructuraProyecto(candidato.FullName))
        {
            return candidato.FullName;
        }

        candidato = candidato.Parent;
    }

    return directorioActual;
}

static bool ExisteEstructuraProyecto(string rutaBase)
{
    return File.Exists(Path.Combine(rutaBase, "SistemaAdministrativoWeb.csproj"))
        && Directory.Exists(Path.Combine(rutaBase, "Views"));
}
