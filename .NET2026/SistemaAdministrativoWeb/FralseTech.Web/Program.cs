using FralseTech.Web.Configuration;
using Microsoft.AspNetCore.Diagnostics.HealthChecks;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.Extensions.Diagnostics.HealthChecks;
using System.Net;

var builder = WebApplication.CreateBuilder(args);

if (builder.Environment.IsDevelopment())
{
    builder.Configuration.AddUserSecrets<Program>(optional: true, reloadOnChange: true);
}

var reverseProxySettings = builder.Configuration
    .GetSection(ReverseProxyOptions.SectionName)
    .Get<ReverseProxyOptions>() ?? new ReverseProxyOptions();
builder.Services.Configure<ReverseProxyOptions>(
    builder.Configuration.GetSection(ReverseProxyOptions.SectionName));

var knownProxyAddresses = new HashSet<IPAddress>();
foreach (var configuredProxy in reverseProxySettings.KnownProxies)
{
    if (!IPAddress.TryParse(configuredProxy?.Trim(), out var proxyAddress))
    {
        throw new InvalidOperationException(
            "ReverseProxy:KnownProxies contiene una dirección IP no válida.");
    }

    knownProxyAddresses.Add(proxyAddress);
}

if (reverseProxySettings.Enabled && knownProxyAddresses.Count == 0)
{
    throw new InvalidOperationException(
        "ReverseProxy está habilitado, pero no tiene ninguna dirección IP válida configurada en KnownProxies.");
}

if (reverseProxySettings.Enabled)
{
    builder.Services.Configure<ForwardedHeadersOptions>(options =>
    {
        options.ForwardedHeaders = ForwardedHeaders.XForwardedFor |
                                   ForwardedHeaders.XForwardedProto;
        options.KnownIPNetworks.Clear();
        options.KnownProxies.Clear();

        foreach (var proxyAddress in knownProxyAddresses)
        {
            options.KnownProxies.Add(proxyAddress);
        }
    });
}

builder.Services.AddControllersWithViews();
builder.Services.AddHealthChecks()
    .AddCheck("self", () => HealthCheckResult.Healthy(), tags: ["live", "ready"]);

var app = builder.Build();

if (reverseProxySettings.Enabled)
{
    app.UseForwardedHeaders();
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseAuthorization();

app.MapStaticAssets();
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}")
    .WithStaticAssets();

app.MapHealthChecks("/healthz/live", new HealthCheckOptions
{
    Predicate = registration => registration.Tags.Contains("live")
});
app.MapHealthChecks("/healthz/ready", new HealthCheckOptions
{
    Predicate = registration => registration.Tags.Contains("ready")
});

app.Run();
