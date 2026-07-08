namespace SistemaAdministrativoWeb.Configuration;

public sealed class MigoApiSettings
{
    public const string SectionName = "MigoApi";

    public string BaseUrl { get; set; } = "https://api.migo.pe/api/v1/";
    public string Token { get; set; } = string.Empty;
    public string ExchangeDatePath { get; set; } = "exchange/date";
    public string ExchangeRangePath { get; set; } = "exchange";
    public string RucPath { get; set; } = "ruc";
    public string DniPath { get; set; } = "dni";
    public string CpePath { get; set; } = "cpe";
}
