using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Localization;
using Microsoft.EntityFrameworkCore;
using SistemaControlEspaciosDeportivosWeb.Data;
using SistemaControlEspaciosDeportivosWeb.Models;
using SistemaControlEspaciosDeportivosWeb.Services;
using System.Globalization;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(connectionString));
builder.Services.AddDatabaseDeveloperPageExceptionFilter();

builder.Services.AddDefaultIdentity<ApplicationUser>(options => options.SignIn.RequireConfirmedAccount = false)
    .AddErrorDescriber<SpanishIdentityErrorDescriber>()
    .AddRoles<IdentityRole>()
    .AddEntityFrameworkStores<ApplicationDbContext>();
builder.Services.AddLocalization(options => options.ResourcesPath = "Resources");
builder.Services.AddControllersWithViews(options =>
    {
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
    options.Cookie.IsEssential = true;
    options.IdleTimeout = TimeSpan.FromMinutes(20);
});
builder.Services.Configure<EmailSettings>(builder.Configuration.GetSection("EmailSettings"));
builder.Services.Configure<AutomationSettings>(builder.Configuration.GetSection("AutomationSettings"));
builder.Services.AddScoped<IModuloPermisoService, ModuloPermisoService>();
builder.Services.AddScoped<ISportCenterStoredProcedureService, SportCenterStoredProcedureService>();
builder.Services.AddScoped<INotificacionEmailService, NotificacionEmailService>();
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
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseRequestLocalization(localizationOptions);
app.UseRouting();
app.UseSession();
app.UseStaticFiles();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

await IdentitySeeder.SeedRolesAsync(app.Services);

app.Run();
