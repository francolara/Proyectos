using GestionServicios.Application.Maintenance;
namespace GestionServicios.Web.ViewModels.Maintenance;

public sealed class TaxRateInput
{
    public string? From { get; set; }
    public string? Value { get; set; }
    public string? Reason { get; set; }
    public string? ExpectedLatest { get; set; }
}
public sealed record TaxRateModel(TaxRatePage Data, TaxRateInput Input, bool CanEdit, string Search, int Page, string? Error = null);
