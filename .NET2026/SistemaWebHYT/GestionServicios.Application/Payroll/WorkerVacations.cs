namespace GestionServicios.Application.Payroll;

public sealed record VacationPeriod(long Id, DateTime From, DateTime To, short Granted,
    int Taken, int Scheduled, int Available);
public sealed record VacationEntry(long Id, long PeriodId, DateTime From, DateTime To,
    short Days, string Status, byte[] Version);
public sealed record WorkerVacations(IReadOnlyList<VacationPeriod> Periods, IReadOnlyList<VacationEntry> Entries);
public sealed record VacationPeriodInput(DateOnly From, DateOnly To, short Days);
public sealed record VacationInput(long PeriodId, DateOnly From, DateOnly To, short Days,
    string Status, long? Id, byte[]? Version);
