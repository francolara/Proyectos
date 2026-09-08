using Microsoft.Extensions.Options;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public sealed class BusinessTimeZoneSettings
{
    public string TimeZoneId { get; set; } = "America/Lima";
}

public interface IBusinessClock
{
    DateTime UtcNow { get; }
    DateTime LocalNow { get; }
    DateOnly Today { get; }
}

public sealed class BusinessClock : IBusinessClock
{
    private static readonly string[] PeruTimeZoneIds = ["America/Lima", "SA Pacific Standard Time"];
    private readonly TimeProvider timeProvider;
    private readonly TimeZoneInfo timeZone;

    public BusinessClock(TimeProvider timeProvider, IOptions<BusinessTimeZoneSettings> options)
    {
        this.timeProvider = timeProvider;
        timeZone = ResolveTimeZone(options.Value.TimeZoneId);
    }

    public DateTime UtcNow => timeProvider.GetUtcNow().UtcDateTime;

    public DateTime LocalNow => TimeZoneInfo.ConvertTime(timeProvider.GetUtcNow(), timeZone).DateTime;

    public DateOnly Today => DateOnly.FromDateTime(LocalNow);

    private static TimeZoneInfo ResolveTimeZone(string? configuredId)
    {
        var candidates = new[] { configuredId }
            .Concat(PeruTimeZoneIds)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase);

        foreach (var candidate in candidates)
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(candidate!);
            }
            catch (TimeZoneNotFoundException)
            {
            }
            catch (InvalidTimeZoneException)
            {
            }
        }

        throw new InvalidOperationException("No se pudo resolver la zona horaria de negocio configurada.");
    }
}
