using System.Net;
using Microsoft.AspNetCore.HttpOverrides;
namespace GestionServicios.Web.Security;
public static class ReverseProxyConfiguration
{
    public static void Configure(WebApplicationBuilder builder)
    {
        var configured=builder.Configuration["ReverseProxy:KnownProxy"];
        if(string.IsNullOrWhiteSpace(configured))
        {
            if(!builder.Environment.IsDevelopment())throw new InvalidOperationException("Configure ReverseProxy:KnownProxy con la IP verificada del proxy de HYT.");
            return;
        }
        if(!IPAddress.TryParse(configured,out var address))throw new InvalidOperationException("ReverseProxy:KnownProxy debe contener una IP válida.");
        builder.Services.Configure<ForwardedHeadersOptions>(options=>
        {
            options.ForwardedHeaders=ForwardedHeaders.XForwardedFor|ForwardedHeaders.XForwardedProto;
            options.ForwardLimit=1;options.RequireHeaderSymmetry=true;options.KnownIPNetworks.Clear();options.KnownProxies.Clear();options.KnownProxies.Add(address);
            if(address.AddressFamily==System.Net.Sockets.AddressFamily.InterNetwork)options.KnownProxies.Add(address.MapToIPv6());
        });
    }
}