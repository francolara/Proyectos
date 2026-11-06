namespace GestionServicios.Contracts;

public sealed record TraceEvent(string Action, DateTimeOffset RecordedAt, string User,
    string? PreviousState = null, string? NewState = null, string? Note = null);

public sealed record RecordTrace(DateTimeOffset CreatedAt, string CreatedBy,
    DateTimeOffset? UpdatedAt, string? UpdatedBy, IReadOnlyList<TraceEvent> Events);
