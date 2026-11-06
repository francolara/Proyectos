using System.Globalization;
using System.Text.Json;
using GestionServicios.Application.Reporting;
namespace GestionServicios.Web.ViewModels.Reporting;
public sealed record DashboardModel(DateOnly From,DateOnly To,JsonElement Data,string? Client=null,string? Store=null,int QueuePage=1);
public sealed record ReportModel(string Kind,DateOnly From,DateOnly To,ReportPage Data,string? Client=null,string? Store=null,string? Situation=null);
public static class ReportingDates
{
    public static readonly TimeZoneInfo Lima=TimeZoneInfo.FindSystemTimeZoneById("America/Lima");
    public static DateOnly Today(TimeProvider clock)=>DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),Lima).DateTime);
    public static DateTimeOffset Start(DateOnly day)=>new(TimeZoneInfo.ConvertTimeToUtc(day.ToDateTime(TimeOnly.MinValue),Lima));
    public static string Number(decimal? value)=>value?.ToString("N2",CultureInfo.GetCultureInfo("es-PE"))??"Pendiente";
}