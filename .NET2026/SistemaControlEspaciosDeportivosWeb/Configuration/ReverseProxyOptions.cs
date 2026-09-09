namespace SistemaControlEspaciosDeportivosWeb.Configuration;

public sealed class ReverseProxyOptions
{
    public const string SectionName = "ReverseProxy";

    public bool Enabled { get; set; }
    public string[] KnownProxies { get; set; } = [];
}
