namespace GestionServicios.Application.Payroll;

public sealed record Worker(long Id, string Document, string Name, string Kind, DateTime StartDate,
    string? Bank, string? MaskedAccount, string? Account, bool Active, byte[] Version, long Total,
    long? CargoId = null, string? CargoName = null, decimal? MonthlySalary = null, WorkerPaySettings? PaySettings = null,
    string? Cci = null, long? PendingVacationDays = null, long[]? Concepts = null, WorkerTelecredit? Telecredit = null, string? GivenNames = null, string? Surnames = null);

public sealed record WorkerInput(long? Id, string Document, string Name, string Kind, DateOnly StartDate,
    bool Active, byte[]? Version, bool UpdateAccount, string? Bank, string? Account,
    long? CargoId = null, decimal? MonthlySalary = null, bool UpdatePosition = false, WorkerPaySettings? PaySettings = null, string? Cci = null, bool UpdateCci = false, long[]? Concepts = null,
    WorkerTelecredit? Telecredit = null, bool UpdateTelecreditAccount = false, string? GivenNames = null, string? Surnames = null);

public sealed record WorkerTelecredit(string? DocumentType, string? Document, string? AccountType, string? Currency);

public sealed record WorkerPaySettings(TimeOnly? Entry, TimeOnly? Exit, string? Pension, string? Afp, string? Frequency);

public sealed record WorkerPosition(long Id, string Name, bool Active, byte[] Version);

public sealed record WorkShift(long Id, DateTime Entry, DateTime? Exit, int ExtraMinutes, byte[] Version,
    long Total, long PeriodExtraMinutes);
public sealed record WorkShiftInput(long? Id, byte[]? Version, DateTime Entry, DateTime? Exit, int ExtraMinutes);

public interface IWorkersRepository
{
    Task<IReadOnlyList<Worker>> WorkersAsync(string actor, long? workerId, string? search, string? kind,
        int page, CancellationToken ct, bool? active = null);
    Task<IReadOnlyList<WorkerPosition>> PositionsAsync(string actor, CancellationToken ct);
    Task<long> SavePositionAsync(string actor, WorkerPosition input, CancellationToken ct);
    Task<long> SaveWorkerAsync(string actor, WorkerInput input, CancellationToken ct);
    Task<IReadOnlyList<WorkShift>> ShiftsAsync(string actor, long workerId, DateOnly? from, DateOnly? to,
        long? shiftId, int page, CancellationToken ct);
    Task<long> SaveShiftAsync(string actor, long workerId, WorkShiftInput input, CancellationToken ct);
    Task<WorkerVacations> VacationsAsync(string actor, long workerId, DateOnly today, CancellationToken ct);
    Task<long> SaveVacationPeriodAsync(string actor, long workerId, VacationPeriodInput input, CancellationToken ct);
    Task<long> SaveVacationAsync(string actor, long workerId, VacationInput input, CancellationToken ct);
}
