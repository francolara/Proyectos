namespace GestionServicios.Application.Payroll;

public sealed record PersonalBalance(string MonedaCodigo, decimal SaldoAnterior, decimal Devengado,
    decimal Adelantado, decimal Descontado, decimal Pagado, decimal Saldo);

public sealed record PersonalMovement(long MovimientoId, string TipoTrabajador, string Concepto,
    DateOnly Fecha, string Descripcion, decimal Importe, string MonedaCodigo, string? Frecuencia,
    DateOnly? PeriodoDesde, DateOnly? PeriodoHasta, string? Referencia, long? CeseId);

public sealed record PersonalTermination(long CeseId, DateOnly FechaCese, DateOnly UltimoDiaTrabajado, string? Motivo, DateOnly? FechaReingreso = null, string? MotivoReingreso = null);

public sealed record PersonalAccount(long TrabajadorId, string NombreCompleto, string Tipo, bool Activo,
    DateOnly Mes, bool Cerrado, PersonalBalance[] Saldos, PersonalMovement[] Movimientos, PersonalTermination[] Ceses);

public sealed record PersonalMovementInput(Guid OperationId, long WorkerId, string Concept, DateOnly Date,
    string Description, decimal Amount, string Currency, string? Frequency, DateOnly? PeriodFrom,
    DateOnly? PeriodTo, string? Reference, long? TerminationId);

public sealed record PersonalTerminationInput(Guid OperationId, long WorkerId, DateOnly Date,
    DateOnly LastWorked, string? Reason);

public sealed record MonthlyStaffBalance(long TrabajadorId, string NombreCompleto, string TipoTrabajador,
    string MonedaCodigo, decimal SaldoAnterior, decimal Devengado, decimal Adelantado, decimal Descontado,
    decimal Pagado, decimal Saldo);

public sealed record MonthlySupplierBalance(long ProveedorId, string RazonSocial, long PresupuestoId,
    string NumeroInterno, string MonedaCodigo, decimal Costo, decimal PagadoMes, decimal PagadoAcumulado, decimal Pendiente);

public sealed record MonthlyPayrollReport(DateOnly Mes, bool VerPersonal, bool VerProveedores, bool Cerrado,
    MonthlyStaffBalance[] Personal, MonthlySupplierBalance[] Proveedores);

public interface IPersonalPaymentsRepository
{
    Task<PersonalAccount?> AccountAsync(string actor, long workerId, DateOnly month, CancellationToken ct);
    Task<long> RegisterMovementAsync(string actor, PersonalMovementInput input, CancellationToken ct);
    Task<long> RegisterReentryAsync(string actor, Guid operationId, long workerId, long terminationId,
        DateOnly date, string? reason, CancellationToken ct);
    Task<long> RegisterTerminationAsync(string actor, PersonalTerminationInput input, CancellationToken ct);
    Task CloseMonthAsync(string actor, DateOnly month, CancellationToken ct);
    Task<MonthlyPayrollReport> ReportAsync(string actor, DateOnly month, long? supplierId, long? budgetId, CancellationToken ct);
}
