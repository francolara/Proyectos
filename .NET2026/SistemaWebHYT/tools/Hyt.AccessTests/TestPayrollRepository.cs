using GestionServicios.Application.Payroll;

sealed class TestPayrollRepository : ISupplierPaymentsRepository, IWorkersRepository,
    IPersonalPaymentsRepository, IPayrollObligationsRepository, IPayPlanningRepository, IPayrollCatalogRepository, IPayrollConceptRepository, ITelecreditRepository
{
    private Supplier supplier = new(1, "20000000001", "Proveedor de prueba", "Banco", "****2345", "000012345", true, new byte[8], 1);
    private SupplierCost? cost;
    private SupplierPayment? payment;
    private Worker worker = new(1, "00000001", "Trabajador de prueba", "PLANILLA",
        new(2025, 1, 1), null, null, null, true, new byte[8], 1);
    private VacationPeriod? period;
    private VacationEntry? vacation;
    private WorkShift? shift;
    private PersonalMovement? movement;
    private PersonalTermination? termination;
    private DateOnly? closedMonth;
    private PayrollObligation? obligation;
    private SctrPolicy? policy;
    private PayPlan? payPlan;
    private readonly List<PayrollConcept> concepts = [
        new(1, "Asignación familiar", "ABONO", "IMPORTE", 100m, null, true, new byte[8]),
        new(2, "Bono", "ABONO", "IMPORTE", null, null, true, new byte[8]),
        new(3, "Sistema de pensiones (AFP)", "CARGO", "PORCENTAJE", null, "AFP", true, new byte[8])
    ];

    private TelecreditAccount? chargeAccount;

    public Task<TelecreditAccount?> GetChargeAccountAsync(string actor, CancellationToken ct)
        => Task.FromResult(chargeAccount);

    public Task<long> SaveChargeAccountAsync(string actor, TelecreditAccount account, CancellationToken ct)
    {
        chargeAccount = account with { VersionFila = new byte[8] };
        return Task.FromResult(1L);
    }

    public Task<TelecreditBatch> ReadBatchAsync(string actor, long planId, byte[] version, string kind, CancellationToken ct)
        => Task.FromResult(new TelecreditBatch(planId, chargeAccount, kind == "PLANILLA"
            ? [new(1, "Pérez Ñahui", "1", "00000042", "BCP", "1910008791097", null, "C", "PEN", 375.10m, false)]
            : [new(2, "López Díaz", "3", "001ABC", "Interbank", null, "00963200000380987111", null, "PEN", 824.90m, false)]));

    public Task<IReadOnlyList<PayrollConcept>> ConceptsAsync(string actor, CancellationToken ct)
        => Task.FromResult<IReadOnlyList<PayrollConcept>>(concepts);

    public Task<long> SaveConceptAsync(string actor, PayrollConcept concept, CancellationToken ct)
    {
        var id = concept.ConceptoId == 0 ? concepts.Max(item => item.ConceptoId) + 1 : concept.ConceptoId;
        concepts.RemoveAll(item => item.ConceptoId == id);
        concepts.Add(concept with { ConceptoId = id, VersionFila = new byte[8] });
        return Task.FromResult(id);
    }

    public Task<PayPlan> PlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct)
        => Task.FromResult(payPlan ?? new PayPlan(null, null, frequency, from, PeriodEnd(frequency, from),
            [new(worker.Id, worker.Name, worker.Kind, 1500m, 60, 0, 0, Banco: "Banco de prueba", Cuenta: "0000042", Cargo: "Supervisor")], VerCuenta: true,
            Conceptos: concepts.Where(item => item.Activo).Select(item =>
                new PayPlanConcept(item.ConceptoId, item.Nombre, item.Naturaleza, item.Calculo, item.Valor)).ToArray()));

    public Task<PayPlan> ReportPlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct)
        => PlanAsync(actor, frequency, from, ct);

    public Task<long> ConfirmPlanAsync(string actor, string frequency, DateOnly from, byte[] version,
        DateOnly date, string reference, CancellationToken ct)
    {
        payPlan = payPlan! with { FechaPago = date, ReferenciaPago = reference };
        return Task.FromResult(payPlan.PlanId!.Value);
    }

    public Task<long> SavePlanAsync(string actor, string frequency, DateOnly from, byte[]? version,
        IReadOnlyList<PayPlanLineInput> lines, CancellationToken ct)
    {
        var row = lines.Single();
        payPlan = new(1, new byte[8], frequency, from, PeriodEnd(frequency, from),
            [new(row.TrabajadorId, worker.Name, worker.Kind, row.Base, row.MinutosExtra, row.ImporteExtra, row.Descuento, row.Ajustes, Cargo: "Supervisor", Incluido: row.Incluido)]);
        return Task.FromResult(1L);
    }

    private static DateOnly PeriodEnd(string frequency, DateOnly from) => frequency switch
    {
        "SEMANAL" => from.AddDays(6),
        "QUINCENAL" when from.Day == 1 => from.AddDays(14),
        _ => new DateOnly(from.Year, from.Month, DateTime.DaysInMonth(from.Year, from.Month))
    };

    public Task<PersonalAccount?> AccountAsync(string actor, long workerId, DateOnly month, CancellationToken ct)
        => Task.FromResult<PersonalAccount?>(new(worker.Id, worker.Name, worker.Kind, termination is null, month,
            closedMonth == month, [new("PEN", 0, 1000, 0, 0, movement?.Importe ?? 0, 1000 - (movement?.Importe ?? 0))],
            movement is null ? [] : [movement], termination is null ? [] : [termination]));

    public Task<long> RegisterMovementAsync(string actor, PersonalMovementInput input, CancellationToken ct)
    {
        movement = new(1, worker.Kind, input.Concept, input.Date, input.Description, input.Amount, input.Currency,
            input.Frequency, input.PeriodFrom, input.PeriodTo, input.Reference, input.TerminationId);
        return Task.FromResult(1L);
    }

    public Task<long> RegisterReentryAsync(string actor, Guid operationId, long workerId, long terminationId,
        DateOnly date, string? reason, CancellationToken ct)
    {
        termination = termination! with { FechaReingreso = date, MotivoReingreso = reason };
        worker = worker with { Active = true, StartDate = date.ToDateTime(TimeOnly.MinValue) };
        return Task.FromResult(terminationId);
    }

    public Task<long> RegisterTerminationAsync(string actor, PersonalTerminationInput input, CancellationToken ct)
    {
        worker = worker with { Active = false };
        termination = new(1, input.Date, input.LastWorked, input.Reason);
        return Task.FromResult(1L);
    }

    public Task CloseMonthAsync(string actor, DateOnly month, CancellationToken ct)
    {
        closedMonth = month;
        return Task.CompletedTask;
    }

    public Task<MonthlyPayrollReport> ReportAsync(string actor, DateOnly month, long? supplierId, long? budgetId, CancellationToken ct)
        => Task.FromResult(new MonthlyPayrollReport(month, true, true, closedMonth == month,
            [new(worker.Id, worker.Name, worker.Kind, "PEN", 0, 1000, 200, 0, 500, 300)],
            [new(supplier.Id, supplier.Name, 1, "PPTO-1", "PEN", 1200, 500, 500, 700)]));

    public Task<PayrollCatalog> CatalogAsync(string actor, CancellationToken ct)
        => Task.FromResult(new PayrollCatalog([new("Interbank", null)], [new("AFP Integra", "AFP"), new("ONP", "ONP")]));

    public Task<PayrollObligations> ObligationsAsync(string actor, DateOnly month, CancellationToken ct)
        => Task.FromResult(new PayrollObligations(month, obligation?.Mes == month ? [obligation] : [],
            policy is null ? [] : [policy], [new(worker.Id, worker.Name), new(2, "Segundo trabajador")]));

    private decimal pensionRate = 10m;

    public Task<IReadOnlyList<PensionEstimate>> PensionEstimatesAsync(string actor, DateOnly month, CancellationToken ct, bool includeWorkers = false)
        => Task.FromResult<IReadOnlyList<PensionEstimate>>([new("Prima AFP", "AFP", pensionRate, 1500m, 15m * pensionRate,
            includeWorkers ? [new(worker.Document, worker.Name, 1500m, 15m * pensionRate)] : null)]);

    public Task<long> SavePensionRateAsync(string actor, string entity, decimal rate, CancellationToken ct)
    {
        pensionRate = rate;
        return Task.FromResult(1L);
    }

    public Task<long> SaveObligationAsync(string actor, PayrollObligation input, CancellationToken ct, bool preservePayment = false)
    {
        obligation = input with { ObligacionId = 1, VersionFila = new byte[8], Estado = input.FechaPago.HasValue ? "PAGADA" : "PENDIENTE" };
        return Task.FromResult(1L);
    }

    public Task<long> SavePolicyAsync(string actor, SctrPolicy input, CancellationToken ct)
    {
        policy = input with
        {
            PolizaId = 1, VersionFila = new byte[8], Estado = "POR_VENCER",
            Trabajadores = input.Trabajadores?.Select(item => item with
                { Nombre = item.TrabajadorId == worker.Id ? worker.Name : "Segundo trabajador" }).ToArray()
        };
        return Task.FromResult(1L);
    }

    public Task<IReadOnlyList<Supplier>> SuppliersAsync(string actor, long? supplierId, string? search, bool activeOnly, int page, CancellationToken ct)
        => Task.FromResult<IReadOnlyList<Supplier>>(supplierId is null ? [supplier with { Account = null }] : [supplier]);

    public Task<long> SaveSupplierAsync(string actor, SupplierInput input, CancellationToken ct)
    {
        supplier = supplier with { Name = input.Name, Account = input.Account };
        return Task.FromResult(1L);
    }

    public Task<IReadOnlyList<SupplierCost>> CostsAsync(string actor, long? budgetId, long? supplierId, long? costId, bool pendingOnly, int page, CancellationToken ct, SupplierCostFilters? filters = null)
        => Task.FromResult<IReadOnlyList<SupplierCost>>(cost is null || (pendingOnly || filters?.Status == "PENDIENTES") && cost.Balance == 0 ? [] : [cost]);

    public Task<long> SaveCostAsync(string actor, SupplierCostInput input, CancellationToken ct)
    {
        cost = new(1, input.BudgetId, "PPTO-1", input.SupplierId, supplier.Name, input.Description, input.Currency,
            input.Document, input.Amount, 0, input.Amount, "PENDIENTE", supplier.Bank, supplier.Account, 1);
        return Task.FromResult(1L);
    }

    public Task<IReadOnlyList<SupplierPayment>> PaymentsAsync(string actor, long costId, int page, CancellationToken ct)
        => Task.FromResult<IReadOnlyList<SupplierPayment>>(payment is null ? [] : [payment]);

    public Task<long> SavePaymentAsync(string actor, SupplierPaymentInput input, CancellationToken ct)
    {
        var remaining = cost!.Balance - input.Amount;
        cost = cost with { Paid = cost.Paid + input.Amount, Balance = remaining, Status = remaining == 0 ? "PAGADO" : "PARCIAL" };
        payment = new(1, input.Amount, input.Date.ToDateTime(TimeOnly.MinValue), input.Reference, cost.Currency, 1,
            input.Frequency, input.PeriodFrom?.ToDateTime(TimeOnly.MinValue), input.PeriodTo?.ToDateTime(TimeOnly.MinValue));
        return Task.FromResult(1L);
    }

    public Task<IReadOnlyList<Worker>> WorkersAsync(string actor, long? workerId, string? search, string? kind, int page, CancellationToken ct, bool? active = null)
        => Task.FromResult<IReadOnlyList<Worker>>(active is null || worker.Active == active ? [worker] : []);

    public Task<IReadOnlyList<WorkerPosition>> PositionsAsync(string actor, CancellationToken ct)
        => Task.FromResult<IReadOnlyList<WorkerPosition>>([new(1, "Supervisor", true, new byte[8])]);

    public Task<long> SavePositionAsync(string actor, WorkerPosition input, CancellationToken ct)
        => Task.FromResult(1L);

    public Task<long> SaveWorkerAsync(string actor, WorkerInput input, CancellationToken ct)
        => Task.FromResult(1L);

    public Task<WorkerVacations> VacationsAsync(string actor, long workerId, DateOnly today, CancellationToken ct)
        => Task.FromResult(new WorkerVacations(period is null ? [] : [period], vacation is null ? [] : [vacation]));

    public Task<long> SaveVacationPeriodAsync(string actor, long workerId, VacationPeriodInput input, CancellationToken ct)
    {
        period = new(1, input.From.ToDateTime(TimeOnly.MinValue), input.To.ToDateTime(TimeOnly.MinValue), input.Days, 0, 0, input.Days);
        return Task.FromResult(1L);
    }

    public Task<long> SaveVacationAsync(string actor, long workerId, VacationInput input, CancellationToken ct)
    {
        vacation = new(1, input.PeriodId, input.From.ToDateTime(TimeOnly.MinValue), input.To.ToDateTime(TimeOnly.MinValue),
            input.Days, input.Status, new byte[8]);
        return Task.FromResult(1L);
    }

    public Task<IReadOnlyList<WorkShift>> ShiftsAsync(string actor, long workerId, DateOnly? from, DateOnly? to, long? shiftId, int page, CancellationToken ct)
        => Task.FromResult<IReadOnlyList<WorkShift>>(shift is null ? [] : [shift]);

    public Task<long> SaveShiftAsync(string actor, long workerId, WorkShiftInput input, CancellationToken ct)
    {
        shift = new(1, input.Entry, input.Exit, input.ExtraMinutes, new byte[8], 1, input.ExtraMinutes);
        return Task.FromResult(1L);
    }
}
