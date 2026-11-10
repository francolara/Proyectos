using System.Data;
using System.Text.Json;
using GestionServicios.Application.Payroll;
using Microsoft.Data.SqlClient;

namespace GestionServicios.Infrastructure.Payroll;

public sealed class SqlSupplierPaymentsRepository(SqlConnectionFactory connections) :
    ISupplierPaymentsRepository, IWorkersRepository, IPersonalPaymentsRepository, IPayrollObligationsRepository, IPayPlanningRepository, IPayrollCatalogRepository
{
    public Task<IReadOnlyList<Worker>> WorkersAsync(string actor, long? workerId, string? search, string? kind, int page, CancellationToken ct, bool? active = null)
        => ReadAsync("spu_TrabajadorHYT_Listar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@Buscar", SqlDbType.NVarChar, search, 180);
            Add(command, "@Tipo", SqlDbType.VarChar, kind, 10);
            if (active.HasValue) Add(command, "@Activo", SqlDbType.Bit, active);
            Add(command, "@Pagina", SqlDbType.Int, page);
        }, reader => new Worker(
            reader.GetFieldValue<long>(reader.GetOrdinal("TrabajadorId")),
            reader.GetFieldValue<string>(reader.GetOrdinal("Documento")),
            reader.GetFieldValue<string>(reader.GetOrdinal("NombreCompleto")),
            reader.GetString(reader.GetOrdinal("Tipo")),
            reader.GetDateTime(reader.GetOrdinal("FechaIngreso")),
            OptionalText(reader, "Banco"),
            OptionalText(reader, "CuentaEnmascarada"),
            OptionalText(reader, "NumeroCuenta"),
            reader.GetFieldValue<bool>(reader.GetOrdinal("Activo")),
            reader.GetFieldValue<byte[]>(reader.GetOrdinal("VersionFila")),
            reader.GetFieldValue<long>(reader.GetOrdinal("Total")),
            reader["CargoId"] as long?, OptionalText(reader, "CargoNombre"), reader["SueldoMensual"] as decimal?,
            new WorkerPaySettings(WorkerTime(reader, "HoraEntradaHabitual"), WorkerTime(reader, "HoraSalidaHabitual"),
                OptionalText(reader, "RegimenPrevisional"), OptionalText(reader, "AdministradoraAFP"),
                OptionalText(reader, "FrecuenciaPago")), OptionalText(reader, "CuentaInterbancaria"),
            reader["VacacionesPendientes"] as long?), ct);

    private static TimeOnly? WorkerTime(SqlDataReader reader, string column)
        => reader[column] is TimeSpan value ? TimeOnly.FromTimeSpan(value) : null;

    public Task<IReadOnlyList<WorkerPosition>> PositionsAsync(string actor, CancellationToken ct)
        => ReadAsync("spu_CargoTrabajadorHYT_Listar", actor, _ => { }, reader =>
            new WorkerPosition(reader.GetInt64(0), reader.GetString(1), reader.GetBoolean(2),
                reader.GetFieldValue<byte[]>(3)), ct);

    public Task<long> SavePositionAsync(string actor, WorkerPosition input, CancellationToken ct)
        => WriteAsync("spu_CargoTrabajadorHYT_Guardar", actor, command =>
        {
            Add(command, "@Nombre", SqlDbType.NVarChar, input.Name, 100);
            Add(command, "@Activo", SqlDbType.Bit, input.Active);
            Add(command, "@CargoId", SqlDbType.BigInt, input.Id == 0 ? null : input.Id);
            Add(command, "@VersionFila", SqlDbType.Binary, input.Id == 0 ? null : input.Version, 8);
        }, ct);

    public Task<long> SaveWorkerAsync(string actor, WorkerInput input, CancellationToken ct)
        => WriteAsync("spu_TrabajadorHYT_Guardar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, input.Id);
            Add(command, "@Documento", SqlDbType.NVarChar, input.Document, 20);
            Add(command, "@NombreCompleto", SqlDbType.NVarChar, input.Name, 180);
            Add(command, "@Tipo", SqlDbType.VarChar, input.Kind, 10);
            Add(command, "@FechaIngreso", SqlDbType.Date, input.StartDate.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Activo", SqlDbType.Bit, input.Active);
            Add(command, "@VersionFila", SqlDbType.Binary, input.Version, 8);
            Add(command, "@ActualizarCuenta", SqlDbType.Bit, input.UpdateAccount);
            Add(command, "@Banco", SqlDbType.NVarChar, input.Bank, 80);
            Add(command, "@NumeroCuenta", SqlDbType.NVarChar, input.Account, 34);
            Add(command, "@CuentaInterbancaria", SqlDbType.VarChar, input.Cci, 100);
            Add(command, "@ActualizarCCI", SqlDbType.Bit, input.UpdateCci);
            Add(command, "@CargoId", SqlDbType.BigInt, input.CargoId);
            Add(command, "@SueldoMensual", SqlDbType.Decimal, input.MonthlySalary);
            Add(command, "@ActualizarPuesto", SqlDbType.Bit, input.UpdatePosition);
            Add(command, "@HoraEntradaHabitual", SqlDbType.Time, input.PaySettings?.Entry?.ToTimeSpan());
            Add(command, "@HoraSalidaHabitual", SqlDbType.Time, input.PaySettings?.Exit?.ToTimeSpan());
            Add(command, "@RegimenPrevisional", SqlDbType.VarChar, input.PaySettings?.Pension, 3);
            Add(command, "@AdministradoraAFP", SqlDbType.NVarChar, input.PaySettings?.Afp, 100);
            Add(command, "@FrecuenciaPago", SqlDbType.VarChar, input.PaySettings?.Frequency, 10);
            Add(command, "@ActualizarPlanificacion", SqlDbType.Bit, input.PaySettings is not null);
        }, ct);

    public Task<IReadOnlyList<Supplier>> SuppliersAsync(string actor, long? supplierId, string? search, bool activeOnly, int page, CancellationToken ct)
        => ReadAsync("spu_ProveedorHYT_Listar", actor, command =>
        {
            Add(command, "@ProveedorId", SqlDbType.BigInt, supplierId);
            Add(command, "@Buscar", SqlDbType.NVarChar, search, 180);
            Add(command, "@SoloActivos", SqlDbType.Bit, activeOnly);
            Add(command, "@Pagina", SqlDbType.Int, page);
        }, reader => new Supplier(
            reader.GetFieldValue<long>(reader.GetOrdinal("ProveedorId")),
            reader.GetFieldValue<string>(reader.GetOrdinal("Ruc")),
            reader.GetFieldValue<string>(reader.GetOrdinal("RazonSocial")),
            OptionalText(reader, "Banco"),
            OptionalText(reader, "CuentaEnmascarada"),
            OptionalText(reader, "NumeroCuenta"),
            reader.GetFieldValue<bool>(reader.GetOrdinal("Activo")),
            reader.GetFieldValue<byte[]>(reader.GetOrdinal("VersionFila")),
            reader.GetFieldValue<long>(reader.GetOrdinal("Total")), OptionalText(reader, "CuentaInterbancaria")), ct);

    public Task<IReadOnlyList<SupplierCost>> CostsAsync(string actor, long? budgetId, long? supplierId, long? costId, bool pendingOnly, int page, CancellationToken ct, SupplierCostFilters? filters = null)
        => ReadAsync("spu_CostoProveedor_Listar", actor, command =>
        {
            Add(command, "@PresupuestoId", SqlDbType.BigInt, budgetId);
            Add(command, "@ProveedorId", SqlDbType.BigInt, supplierId);
            Add(command, "@CostoProveedorId", SqlDbType.BigInt, costId);
            Add(command, "@SoloPendientes", SqlDbType.Bit, pendingOnly);
            Add(command, "@Pagina", SqlDbType.Int, page);
            Add(command, "@Desde", SqlDbType.Date, filters?.From?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Hasta", SqlDbType.Date, filters?.To?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@EstadoFiltro", SqlDbType.VarChar, filters?.Status, 10);
            Add(command, "@Buscar", SqlDbType.NVarChar, filters?.Search, 180);
        }, reader => new SupplierCost(
            reader.GetFieldValue<long>(reader.GetOrdinal("CostoProveedorId")),
            reader.GetFieldValue<long>(reader.GetOrdinal("PresupuestoId")),
            reader.GetFieldValue<string>(reader.GetOrdinal("NumeroInterno")),
            reader.GetFieldValue<long>(reader.GetOrdinal("ProveedorId")),
            reader.GetFieldValue<string>(reader.GetOrdinal("RazonSocial")),
            reader.GetFieldValue<string>(reader.GetOrdinal("Descripcion")),
            reader.GetFieldValue<string>(reader.GetOrdinal("MonedaCodigo")),
            OptionalText(reader, "DocumentoSustento"),
            reader.GetFieldValue<decimal>(reader.GetOrdinal("Importe")),
            reader.GetFieldValue<decimal>(reader.GetOrdinal("Pagado")),
            reader.GetFieldValue<decimal>(reader.GetOrdinal("Pendiente")),
            reader.GetFieldValue<string>(reader.GetOrdinal("Estado")),
            OptionalText(reader, "Banco"),
            OptionalText(reader, "NumeroCuenta"),
            reader.GetFieldValue<long>(reader.GetOrdinal("Total")), OptionalDate(reader, "FechaCosto")), ct);

    public Task<IReadOnlyList<SupplierPayment>> PaymentsAsync(string actor, long costId, int page, CancellationToken ct)
        => ReadAsync("spu_PagoProveedor_Listar", actor, command =>
        {
            Add(command, "@CostoProveedorId", SqlDbType.BigInt, costId);
            Add(command, "@Pagina", SqlDbType.Int, page);
        }, reader => new SupplierPayment(
            reader.GetFieldValue<long>(reader.GetOrdinal("PagoProveedorId")),
            reader.GetFieldValue<decimal>(reader.GetOrdinal("Importe")),
            reader.GetFieldValue<DateTime>(reader.GetOrdinal("FechaPago")),
            reader.GetFieldValue<string>(reader.GetOrdinal("Referencia")),
            reader.GetFieldValue<string>(reader.GetOrdinal("MonedaCodigo")),
            reader.GetFieldValue<long>(reader.GetOrdinal("Total")),
            OptionalText(reader, "Frecuencia"), OptionalDate(reader, "PeriodoDesde"),
            OptionalDate(reader, "PeriodoHasta")), ct);

    public Task<long> SaveSupplierAsync(string actor, SupplierInput input, CancellationToken ct)
        => WriteAsync("spu_ProveedorHYT_Guardar", actor, command =>
        {
            Add(command, "@ProveedorId", SqlDbType.BigInt, input.Id);
            Add(command, "@Ruc", SqlDbType.VarChar, input.TaxId, 11);
            Add(command, "@RazonSocial", SqlDbType.NVarChar, input.Name, 180);
            Add(command, "@Activo", SqlDbType.Bit, input.Active);
            Add(command, "@VersionFila", SqlDbType.Binary, input.Version, 8);
            Add(command, "@ActualizarCuenta", SqlDbType.Bit, input.UpdateAccount);
            Add(command, "@Banco", SqlDbType.NVarChar, input.Bank, 80);
            Add(command, "@NumeroCuenta", SqlDbType.NVarChar, input.Account, 34);
            Add(command, "@CuentaInterbancaria", SqlDbType.VarChar, input.Cci, 100);
            Add(command, "@ActualizarCCI", SqlDbType.Bit, input.UpdateCci);
        }, ct);

    public Task<long> SaveCostAsync(string actor, SupplierCostInput input, CancellationToken ct)
        => WriteAsync("spu_CostoProveedor_Registrar", actor, command =>
        {
            Add(command, "@OperacionId", SqlDbType.UniqueIdentifier, input.OperationId);
            Add(command, "@PresupuestoId", SqlDbType.BigInt, input.BudgetId);
            Add(command, "@ProveedorId", SqlDbType.BigInt, input.SupplierId);
            Add(command, "@Descripcion", SqlDbType.NVarChar, input.Description, 300);
            Add(command, "@Importe", SqlDbType.Decimal, input.Amount);
            Add(command, "@MonedaCodigo", SqlDbType.NVarChar, input.Currency, 3);
            Add(command, "@DocumentoSustento", SqlDbType.NVarChar, input.Document, 100);
        }, ct);

    public Task<long> SavePaymentAsync(string actor, SupplierPaymentInput input, CancellationToken ct)
        => WriteAsync("spu_PagoProveedor_Registrar", actor, command =>
        {
            Add(command, "@OperacionId", SqlDbType.UniqueIdentifier, input.OperationId);
            Add(command, "@CostoProveedorId", SqlDbType.BigInt, input.CostId);
            Add(command, "@Importe", SqlDbType.Decimal, input.Amount);
            Add(command, "@FechaPago", SqlDbType.Date, input.Date.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Referencia", SqlDbType.NVarChar, input.Reference, 100);
            Add(command, "@Frecuencia", SqlDbType.VarChar, input.Frequency, 10);
            Add(command, "@PeriodoDesde", SqlDbType.Date, input.PeriodFrom?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@PeriodoHasta", SqlDbType.Date, input.PeriodTo?.ToDateTime(TimeOnly.MinValue));
        }, ct);

    public Task<IReadOnlyList<WorkShift>> ShiftsAsync(string actor, long workerId, DateOnly? from, DateOnly? to,
        long? shiftId, int page, CancellationToken ct)
        => ReadAsync("spu_JornadaHYT_Listar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@Desde", SqlDbType.Date, from?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Hasta", SqlDbType.Date, to?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@JornadaId", SqlDbType.BigInt, shiftId);
            Add(command, "@Pagina", SqlDbType.Int, page);
        }, reader =>
        {
            var exit = reader.GetOrdinal("Salida");
            return new WorkShift(reader.GetInt64(reader.GetOrdinal("JornadaId")),
                reader.GetDateTime(reader.GetOrdinal("Entrada")),
                reader.IsDBNull(exit) ? null : reader.GetDateTime(exit),
                reader.GetInt32(reader.GetOrdinal("MinutosExtrasAprobados")),
                reader.GetFieldValue<byte[]>(reader.GetOrdinal("VersionFila")),
                reader.GetInt64(reader.GetOrdinal("Total")), reader.GetInt64(reader.GetOrdinal("MinutosExtrasTotal")));
        }, ct);

    public Task<long> SaveShiftAsync(string actor, long workerId, WorkShiftInput input, CancellationToken ct)
        => WriteAsync("spu_JornadaHYT_Guardar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@Entrada", SqlDbType.DateTime2, input.Entry);
            Add(command, "@Salida", SqlDbType.DateTime2, input.Exit);
            Add(command, "@MinutosExtrasAprobados", SqlDbType.Int, input.ExtraMinutes);
            Add(command, "@JornadaId", SqlDbType.BigInt, input.Id);
            Add(command, "@VersionFila", SqlDbType.Binary, input.Version, 8);
        }, ct);

    public async Task<WorkerVacations> VacationsAsync(string actor, long workerId, DateOnly today, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = CreateCommand(connection, "spu_VacacionesHYT_Consultar", actor, sql =>
        {
            Add(sql, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(sql, "@Hoy", SqlDbType.Date, today.ToDateTime(TimeOnly.MinValue));
        });
        await using var reader = await command.ExecuteReaderAsync(ct);
        var periods = new List<VacationPeriod>();
        while (await reader.ReadAsync(ct))
        {
            periods.Add(new(reader.GetInt64(reader.GetOrdinal("PeriodoVacacionesId")),
                reader.GetDateTime(reader.GetOrdinal("FechaDesde")), reader.GetDateTime(reader.GetOrdinal("FechaHasta")),
                reader.GetInt16(reader.GetOrdinal("DiasReconocidos")), reader.GetInt32(reader.GetOrdinal("DiasGozados")),
                reader.GetInt32(reader.GetOrdinal("DiasProgramados")), reader.GetInt32(reader.GetOrdinal("DiasDisponibles"))));
        }

        var entries = new List<VacationEntry>();
        if (await reader.NextResultAsync(ct))
        {
            while (await reader.ReadAsync(ct))
            {
                entries.Add(new(reader.GetInt64(reader.GetOrdinal("VacacionesId")),
                    reader.GetInt64(reader.GetOrdinal("PeriodoVacacionesId")),
                    reader.GetDateTime(reader.GetOrdinal("FechaInicio")), reader.GetDateTime(reader.GetOrdinal("FechaFin")),
                    reader.GetInt16(reader.GetOrdinal("Dias")), reader.GetString(reader.GetOrdinal("Estado")),
                    reader.GetFieldValue<byte[]>(reader.GetOrdinal("VersionFila"))));
            }
        }
        return new(periods, entries);
    }

    public Task<long> SaveVacationPeriodAsync(string actor, long workerId, VacationPeriodInput input, CancellationToken ct)
        => WriteAsync("spu_PeriodoVacacionesHYT_Registrar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@FechaDesde", SqlDbType.Date, input.From.ToDateTime(TimeOnly.MinValue));
            Add(command, "@FechaHasta", SqlDbType.Date, input.To.ToDateTime(TimeOnly.MinValue));
            Add(command, "@DiasReconocidos", SqlDbType.SmallInt, input.Days);
        }, ct);

    public Task<long> SaveVacationAsync(string actor, long workerId, VacationInput input, CancellationToken ct)
        => WriteAsync("spu_VacacionesHYT_Guardar", actor, command =>
        {
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@PeriodoVacacionesId", SqlDbType.BigInt, input.PeriodId);
            Add(command, "@FechaInicio", SqlDbType.Date, input.From.ToDateTime(TimeOnly.MinValue));
            Add(command, "@FechaFin", SqlDbType.Date, input.To.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Dias", SqlDbType.SmallInt, input.Days);
            Add(command, "@Estado", SqlDbType.VarChar, input.Status, 10);
            Add(command, "@VacacionesId", SqlDbType.BigInt, input.Id);
            Add(command, "@VersionFila", SqlDbType.Binary, input.Version, 8);
        }, ct);

    public async Task<PersonalAccount?> AccountAsync(string actor, long workerId, DateOnly month, CancellationToken ct)
    {
        var result = await ReadJsonAsync<PersonalAccount>("spu_EstadoCuentaPersonalHYT_Consultar", actor, sql =>
        {
            Add(sql, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(sql, "@Mes", SqlDbType.Date, month.ToDateTime(TimeOnly.MinValue));
        }, ct);
        return result is null ? null : result with
        {
            Saldos = result.Saldos ?? [], Movimientos = result.Movimientos ?? [], Ceses = result.Ceses ?? []
        };
    }

    public async Task<MonthlyPayrollReport> ReportAsync(string actor, DateOnly month, long? supplierId,
        long? budgetId, CancellationToken ct)
    {
        var result = await ReadJsonAsync<MonthlyPayrollReport>("spu_ReportePlanillasHYT_Consultar", actor, sql =>
        {
            Add(sql, "@Mes", SqlDbType.Date, month.ToDateTime(TimeOnly.MinValue));
            Add(sql, "@ProveedorId", SqlDbType.BigInt, supplierId);
            Add(sql, "@PresupuestoId", SqlDbType.BigInt, budgetId);
        }, ct) ?? throw new InvalidOperationException("SQL no devolvió el reporte de planillas.");
        return result with { Personal = result.Personal ?? [], Proveedores = result.Proveedores ?? [] };
    }

    public async Task<PayrollCatalog> CatalogAsync(string actor, CancellationToken ct)
        => await ReadJsonAsync<PayrollCatalog>("spu_CatalogosPlanillasHYT_Consultar", actor, _ => { }, ct)
            ?? new([], []);

    public async Task<PayrollObligations> ObligationsAsync(string actor, DateOnly month, CancellationToken ct)
    {
        var result = await ReadJsonAsync<PayrollObligations>("spu_ObligacionesHYT_Consultar", actor,
            sql => Add(sql, "@Mes", SqlDbType.Date, month.ToDateTime(TimeOnly.MinValue)), ct)
            ?? throw new InvalidOperationException("SQL no devolvió las obligaciones del mes.");
        return result with { Obligaciones = result.Obligaciones ?? [], Polizas = result.Polizas ?? [] };
    }

    public Task<IReadOnlyList<PensionEstimate>> PensionEstimatesAsync(string actor, DateOnly month, CancellationToken ct, bool includeWorkers = false)
        => ReadAsync("spu_EstimadoPrevisionalHYT_Consultar", actor,
            command =>
            {
                Add(command, "@Mes", SqlDbType.Date, month.ToDateTime(TimeOnly.MinValue));
                if (includeWorkers) Add(command, "@IncluirTrabajadores", SqlDbType.Bit, true);
            },
            reader => new PensionEstimate(reader.GetString(0), reader.GetString(1), reader.GetDecimal(2),
                reader.GetDecimal(3), reader.GetDecimal(4),
                includeWorkers ? JsonSerializer.Deserialize<PensionEstimateWorker[]>(reader.GetString(5)) : null), ct);

    public Task<long> SavePensionRateAsync(string actor, string entity, decimal rate, CancellationToken ct)
        => WriteAsync("spu_EntidadPrevisionalHYT_GuardarTasa", actor, command =>
        {
            Add(command, "@Entidad", SqlDbType.NVarChar, entity, 100);
            Add(command, "@Porcentaje", SqlDbType.Decimal, rate);
        }, ct);

    public Task<long> SaveObligationAsync(string actor, PayrollObligation input, CancellationToken ct, bool preservePayment = false)
        => WriteAsync("spu_ObligacionPersonalHYT_Guardar", actor, command =>
        {
            Add(command, "@Tipo", SqlDbType.VarChar, input.Tipo, 3);
            Add(command, "@Entidad", SqlDbType.NVarChar, input.Entidad, 120);
            Add(command, "@Mes", SqlDbType.Date, input.Mes.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Importe", SqlDbType.Decimal, input.Importe);
            Add(command, "@Vencimiento", SqlDbType.Date, input.Vencimiento.ToDateTime(TimeOnly.MinValue));
            Add(command, "@FechaPago", SqlDbType.Date, input.FechaPago?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@ReferenciaPago", SqlDbType.NVarChar, input.ReferenciaPago, 100);
            if (preservePayment) Add(command, "@ConservarPago", SqlDbType.Bit, true);
            Add(command, "@ObligacionId", SqlDbType.BigInt, input.ObligacionId == 0 ? null : input.ObligacionId);
            Add(command, "@VersionFila", SqlDbType.Binary, input.ObligacionId == 0 ? null : input.VersionFila, 8);
        }, ct);

    public Task<long> SavePolicyAsync(string actor, SctrPolicy input, CancellationToken ct)
        => WriteAsync("spu_PolizaSctrHYT_Guardar", actor, command =>
        {
            Add(command, "@Aseguradora", SqlDbType.NVarChar, input.Aseguradora, 120);
            Add(command, "@NumeroPoliza", SqlDbType.NVarChar, input.NumeroPoliza, 80);
            Add(command, "@Cobertura", SqlDbType.VarChar, input.Cobertura, 10);
            Add(command, "@VigenciaDesde", SqlDbType.Date, input.VigenciaDesde.ToDateTime(TimeOnly.MinValue));
            Add(command, "@VigenciaHasta", SqlDbType.Date, input.VigenciaHasta.ToDateTime(TimeOnly.MinValue));
            Add(command, "@PolizaId", SqlDbType.BigInt, input.PolizaId == 0 ? null : input.PolizaId);
            Add(command, "@VersionFila", SqlDbType.Binary, input.PolizaId == 0 ? null : input.VersionFila, 8);
            Add(command, "@Trabajadores", SqlDbType.NVarChar,
                input.Trabajadores is null ? null : JsonSerializer.Serialize(input.Trabajadores.Select(worker => worker.TrabajadorId)), -1);
        }, ct);

    public Task<PayPlan> PlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct)
        => ReadPlanAsync(actor, frequency, from, false, ct);

    public Task<PayPlan> ReportPlanAsync(string actor, string frequency, DateOnly from, CancellationToken ct)
        => ReadPlanAsync(actor, frequency, from, true, ct);

    private async Task<PayPlan> ReadPlanAsync(string actor, string frequency, DateOnly from, bool report, CancellationToken ct)
        => await ReadJsonAsync<PayPlan>("spu_PlanPagoHYT_Consultar", actor, command =>
        {
            Add(command, "@Frecuencia", SqlDbType.VarChar, frequency, 10);
            Add(command, "@Desde", SqlDbType.Date, from.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Reporte", SqlDbType.Bit, report);
        }, ct) ?? throw new InvalidOperationException("SQL no devolvió la planificación.");

    public Task<long> ConfirmPlanAsync(string actor, string frequency, DateOnly from, byte[] version,
        DateOnly date, string reference, CancellationToken ct)
        => WriteAsync("spu_PlanPagoHYT_ConfirmarPago", actor, command =>
        {
            Add(command, "@Frecuencia", SqlDbType.VarChar, frequency, 10);
            Add(command, "@Desde", SqlDbType.Date, from.ToDateTime(TimeOnly.MinValue));
            Add(command, "@VersionFila", SqlDbType.Binary, version, 8);
            Add(command, "@FechaPago", SqlDbType.Date, date.ToDateTime(TimeOnly.MinValue));
            Add(command, "@ReferenciaPago", SqlDbType.NVarChar, reference, 100);
        }, ct);

    public Task<long> SavePlanAsync(string actor, string frequency, DateOnly from, byte[]? version,
        IReadOnlyList<PayPlanLineInput> lines, CancellationToken ct)
        => WriteAsync("spu_PlanPagoHYT_Guardar", actor, command =>
        {
            Add(command, "@Frecuencia", SqlDbType.VarChar, frequency, 10);
            Add(command, "@Desde", SqlDbType.Date, from.ToDateTime(TimeOnly.MinValue));
            Add(command, "@VersionFila", SqlDbType.Binary, version, 8);
            Add(command, "@Filas", SqlDbType.NVarChar, JsonSerializer.Serialize(lines), -1);
        }, ct);

    private async Task<T?> ReadJsonAsync<T>(string procedure, string actor, Action<SqlCommand> bind, CancellationToken ct)
        where T : class
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = CreateCommand(connection, procedure, actor, bind);
        var result = await command.ExecuteScalarAsync(ct);
        return result is string json && !string.IsNullOrWhiteSpace(json)
            ? JsonSerializer.Deserialize<T>(json)
            : null;
    }

    public Task<long> RegisterMovementAsync(string actor, PersonalMovementInput input, CancellationToken ct)
        => WriteAsync("spu_MovimientoPersonalHYT_Registrar", actor, command =>
        {
            Add(command, "@OperacionId", SqlDbType.UniqueIdentifier, input.OperationId);
            Add(command, "@TrabajadorId", SqlDbType.BigInt, input.WorkerId);
            Add(command, "@Concepto", SqlDbType.VarChar, input.Concept, 24);
            Add(command, "@Fecha", SqlDbType.Date, input.Date.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Descripcion", SqlDbType.NVarChar, input.Description, 300);
            Add(command, "@Importe", SqlDbType.Decimal, input.Amount);
            Add(command, "@MonedaCodigo", SqlDbType.NVarChar, input.Currency, 3);
            Add(command, "@Frecuencia", SqlDbType.VarChar, input.Frequency, 10);
            Add(command, "@PeriodoDesde", SqlDbType.Date, input.PeriodFrom?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@PeriodoHasta", SqlDbType.Date, input.PeriodTo?.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Referencia", SqlDbType.NVarChar, input.Reference, 100);
            Add(command, "@CeseId", SqlDbType.BigInt, input.TerminationId);
        }, ct);

    public Task<long> RegisterReentryAsync(string actor, Guid operationId, long workerId, long terminationId,
        DateOnly date, string? reason, CancellationToken ct)
        => WriteAsync("spu_CeseTrabajadorHYT_Reingresar", actor, command =>
        {
            Add(command, "@OperacionId", SqlDbType.UniqueIdentifier, operationId);
            Add(command, "@TrabajadorId", SqlDbType.BigInt, workerId);
            Add(command, "@CeseId", SqlDbType.BigInt, terminationId);
            Add(command, "@FechaReingreso", SqlDbType.Date, date.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Motivo", SqlDbType.NVarChar, reason, 300);
        }, ct);

    public Task<long> RegisterTerminationAsync(string actor, PersonalTerminationInput input, CancellationToken ct)
        => WriteAsync("spu_CeseTrabajadorHYT_Registrar", actor, command =>
        {
            Add(command, "@OperacionId", SqlDbType.UniqueIdentifier, input.OperationId);
            Add(command, "@TrabajadorId", SqlDbType.BigInt, input.WorkerId);
            Add(command, "@FechaCese", SqlDbType.Date, input.Date.ToDateTime(TimeOnly.MinValue));
            Add(command, "@UltimoDiaTrabajado", SqlDbType.Date, input.LastWorked.ToDateTime(TimeOnly.MinValue));
            Add(command, "@Motivo", SqlDbType.NVarChar, input.Reason, 300);
        }, ct);

    public async Task CloseMonthAsync(string actor, DateOnly month, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = CreateCommand(connection, "spu_CierrePersonalHYT_Registrar", actor,
            sql => Add(sql, "@Mes", SqlDbType.Date, month.ToDateTime(TimeOnly.MinValue)));
        await command.ExecuteNonQueryAsync(ct);
    }

    private async Task<IReadOnlyList<T>> ReadAsync<T>(string procedure, string actor,
        Action<SqlCommand> bind, Func<SqlDataReader, T> map, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = CreateCommand(connection, procedure, actor, bind);
        await using var reader = await command.ExecuteReaderAsync(ct);
        var items = new List<T>();
        while (await reader.ReadAsync(ct))
            items.Add(map(reader));
        return items;
    }

    private async Task<long> WriteAsync(string procedure, string actor, Action<SqlCommand> bind, CancellationToken ct)
    {
        await using var connection = connections.CreateConnection();
        await connection.OpenAsync(ct);
        await using var command = CreateCommand(connection, procedure, actor, bind);
        var result = await command.ExecuteScalarAsync(ct);
        return result is long id ? id : throw new InvalidOperationException("SQL no devolvió el identificador del registro.");
    }

    private static SqlCommand CreateCommand(SqlConnection connection, string procedure, string actor, Action<SqlCommand> bind)
    {
        var command = new SqlCommand($"dbo.{procedure}", connection) { CommandType = CommandType.StoredProcedure };
        Add(command, "@UsuarioId", SqlDbType.NVarChar, actor, 128);
        bind(command);
        return command;
    }

    private static void Add(SqlCommand command, string name, SqlDbType type, object? value, int size = 0)
    {
        var parameter = command.Parameters.Add(name, type, size);
        if (type == SqlDbType.Decimal)
        {
            parameter.Precision = 20;
            parameter.Scale = 4;
        }
        parameter.Value = value ?? DBNull.Value;
    }

    private static string? OptionalText(SqlDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);
    }

    private static DateTime? OptionalDate(SqlDataReader reader, string column)
    {
        var ordinal = reader.GetOrdinal(column);
        return reader.IsDBNull(ordinal) ? null : reader.GetDateTime(ordinal);
    }
}
