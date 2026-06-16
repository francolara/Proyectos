using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging.EventLog;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.Data;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;

var builder = WebApplication.CreateBuilder(args);
builder.Logging.AddFilter<EventLogLoggerProvider>(level => level >= LogLevel.None);

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

builder.Services.AddDefaultIdentity<IdentityUser>(options =>
    {
        options.SignIn.RequireConfirmedAccount = false;
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
builder.Services.AddHttpContextAccessor();
builder.Services.AddSession(options =>
{
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
    options.IdleTimeout = TimeSpan.FromHours(8);
});

builder.Services.AddScoped<IDbConnectionFactory, SqlConnectionFactory>();
builder.Services.AddScoped<IEmpresaRepository, EmpresaRepository>();
builder.Services.AddScoped<ICurrentCompanyAccessor, SessionCurrentCompanyAccessor>();
builder.Services.AddScoped<ICuentaAdministradoraRepository, CuentaAdministradoraRepository>();
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
