using System.Globalization;
using System.Security.Claims;
using System.Text.Json;
using System.Xml;
using GestionServicios.Application.Billing;
using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using GestionServicios.Web.Security;
using GestionServicios.Web.ViewModels.Billing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace GestionServicios.Web.Controllers;

// Firma: FRANCO LARA - 10/11/2026 | Compras y comprobantes, carga UBL individual o atómica por lote, registro manual con totales calculados, ayudas, pagos y vínculos desde ficha y listado.
// Firma: FRANCO LARA - 11/11/2026 | Filtra las ventas por presencia de factoring y consulta la detracción de las partidas y permite buscar y asociar órdenes de compra y buscar trabajos disponibles por factura y presupuestos del cliente para precargar y facturas emitidas para notas de crédito.
[Authorize]
[Route("Comprobantes/{scope:regex(^(compras|ventas)$)}")]
public sealed class CommercialDocumentsController(ICommercialDocumentRepository documents, IBudgetRepository budgets,
    TimeProvider clock) : Controller
{
    private string Actor => User.FindFirstValue(ClaimTypes.NameIdentifier)!;
    private DateOnly Today => DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(clock.GetUtcNow(),
        TimeZoneInfo.FindSystemTimeZoneById("America/Lima")).DateTime);
    private bool Has(string permission) => User.HasClaim(AccessService.PermissionClaim, permission);
    private bool Allowed(string scope, string action) => Has((scope == "ventas" ? "FACTURADO." : "PAGOS_PROVEEDOR.") + action);

    [HttpGet("")]
    public async Task<IActionResult> Index(string scope, DateOnly? desde, DateOnly? hasta, string? tipo, string? buscar,
        int pagina = 1, bool sinFactoring = false, bool? factoring = null, CancellationToken ct = default)
    {
        if (!Allowed(scope, "CONSULTAR")) return Forbid();
        var end = hasta ?? Today;
        var start = desde ?? end.AddDays(-30);
        if (!ModelState.IsValid || start > end || pagina is < 1 or > 100000 || buscar?.Length > 150 ||
            tipo is not (null or "01" or "03" or "07")) return BadRequest();
        var sales = scope == "ventas";
        var hasFactoring = sales ? factoring : null;
        var onlyWithoutFactoring = sales && hasFactoring is null && sinFactoring;
        var data = await documents.ReadAsync(Actor, sales, null, start, end, tipo, buscar, pagina, null, null, ct,
            onlyWithoutFactoring, hasFactoring);
        return View(new CommercialPage(scope, data, Allowed(scope, "EDITAR"), Today, start, end, tipo, buscar, pagina,
            onlyWithoutFactoring, hasFactoring));
    }

    [HttpGet("{id:long}")]
    public async Task<IActionResult> Details(string scope, long id, bool print, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR")) return Forbid();
        var data = await ReadDocumentAsync(scope == "ventas", id, ct);
        if (data is null) return NotFound();
        ViewData["IsPrint"] = print;
        return View(new CommercialPage(scope, data.Value, !print && Allowed(scope, "EDITAR"), Today));
    }

    [HttpGet("{id:long}/Vincular")]
    public async Task<IActionResult> LinkForm(string scope, long id, bool popup, string? retorno, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR") ||
            scope == "compras" && (!Has("COSTOS_PROVEEDOR.CONSULTAR") || !Has("COSTOS_PROVEEDOR.EDITAR"))) return Forbid();
        if (!popup) return RedirectToAction(nameof(Details), new { scope, id });
        var data = await ReadDocumentAsync(scope == "ventas", id, ct);
        if (data is null) return NotFound();
        return PartialView("_LinkSelection", (new CommercialPage(scope, data.Value, true, Today), true, LocalReturn(scope, retorno)));
    }

    [HttpGet("{id:long}/TrabajosDisponibles")]
    public async Task<IActionResult> AvailableSupplierWork(string scope, long id, string? buscar, CancellationToken ct)
    {
        if (scope != "compras" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR") ||
            !Has("COSTOS_PROVEEDOR.CONSULTAR") || !Has("COSTOS_PROVEEDOR.EDITAR")) return Forbid();
        if (id <= 0 || buscar?.Length > 150) return BadRequest();
        var data = await ReadDocumentAsync(false, id, ct, buscar?.Trim());
        return data is null ? NotFound() : PartialView("_LinkOptions", new CommercialPage(scope, data.Value, true, Today));
    }

    private async Task<JsonElement?> ReadDocumentAsync(bool sales, long id, CancellationToken ct, string? costSearch = null)
    {
        var data = await documents.ReadAsync(Actor, sales, id, null, null, null, costSearch, 1, null, null, ct);
        if (!data.TryGetProperty("Documento", out var record) || record.ValueKind != JsonValueKind.Object) return null;
        return sales
            ? await documents.ReadAsync(Actor, true, id, null, null, null, null, 1, record.GetProperty("ClienteId").GetInt64(), null, ct)
            : data;
    }

    private string LocalReturn(string scope, string? target)
        => Url.IsLocalUrl(target) ? target! : Url.Action(nameof(Index), new { scope })!;

    [HttpGet("Nuevo")]
    public async Task<IActionResult> New(string scope, long? clienteId, long? presupuestoId, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        var sales = scope == "ventas";
        var input = new CommercialInput { Fecha = Today };
        var data = await documents.ReadAsync(Actor, sales, null, input.Fecha, null, null, null, 1, clienteId, presupuestoId, ct);
        if (sales && data.TryGetProperty("Cliente", out var client) && client.ValueKind == JsonValueKind.Object)
        {
            input.DocumentoTercero = CommercialPage.Text(client, "DocumentoTercero");
            input.NombreTercero = CommercialPage.Text(client, "NombreTercero");
            if (presupuestoId is > 0 && Has("PRESUPUESTOS.CONSULTAR") &&
                data.GetProperty("OpcionesPresupuesto").EnumerateArray().Any(row => row.GetProperty("PresupuestoId").GetInt64() == presupuestoId))
            {
                var budget = await budgets.GetAsync(Actor, presupuestoId.Value, null, 1, ct);
                if (budget is not null)
                {
                    input.Moneda = BudgetRecord.Text(budget.Revision, "CodigoMoneda");
                    var includedTax = BudgetRecord.Text(budget.Revision, "CodigoTipoIGV") == "CON_IGV";
                    var divisor = includedTax ? 1 + CommercialPage.Amount(budget.Revision, "TasaIGVAplicada") : 1;
                    input.Presupuestos = [presupuestoId.Value];
                    input.Detalle = budget.Data.GetProperty("Partidas").EnumerateArray().Select(line => new CommercialLine(
                        BudgetRecord.Text(line,"CodigoInterno"), BudgetRecord.Text(line,"Descripcion"), BudgetRecord.Text(line,"Unidad"),
                        CommercialPage.Amount(line,"Cantidad"), BudgetAmounts.Round(CommercialPage.Amount(line,"PrecioUnitario") / divisor),
                        CommercialPage.Amount(line,"BaseImponible"))).ToList();
                }
            }
        }
        if (input.Detalle.Count == 0) input.Detalle.Add(new(null, "", "", 1, 0, 0));
        return View(new CommercialForm(scope, data, input, clienteId, input.Presupuestos.Count > 0 ? presupuestoId : null, Today));
    }

    [HttpGet("Presupuestos")]
    public async Task<IActionResult> SalesBudgets(string scope, long clienteId, string? buscar, CancellationToken ct)
    {
        if (scope != "ventas" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR") ||
            !Has("PRESUPUESTOS.CONSULTAR")) return Forbid();
        if (!ModelState.IsValid || clienteId <= 0 || buscar?.Length > 150) return BadRequest();
        var data = await documents.ReadAsync(Actor, true, null, null, null, null, buscar?.Trim(), 1, clienteId, null, ct);
        return Json(data.GetProperty("OpcionesPresupuesto"));
    }

    [HttpGet("Terceros")]
    public async Task<IActionResult> Parties(string scope, string? buscar, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (buscar?.Length > 150) return BadRequest();
        var sales = scope == "ventas";
        var search = buscar?.Trim() ?? "";
        var data = await documents.ReadAsync(Actor, sales, null, null, null, null, search, 1, null, null, ct);
        return Json(data.GetProperty(sales ? "OpcionesCliente" : "OpcionesProveedor"));
    }

    [HttpGet("{id:long}/Presupuestos")]
    public async Task<IActionResult> CostBudgets(string scope, long id, string? buscar, CancellationToken ct)
    {
        if (scope != "compras" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR") ||
            !Has("COSTOS_PROVEEDOR.EDITAR") || !Has("PRESUPUESTOS.CONSULTAR")) return Forbid();
        if (buscar?.Length > 150) return BadRequest();
        var data = await documents.ReadAsync(Actor, false, id, null, null, null, buscar?.Trim(), 1, null, null, ct);
        return Json(data.GetProperty("OpcionesCostoPresupuesto"));
    }

    [HttpGet("FacturasEmitidas")]
    public async Task<IActionResult> IssuedInvoices(string scope, long clienteId, DateOnly? desde, DateOnly? hasta,
        string? buscar, CancellationToken ct)
    {
        if (scope != "ventas" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        var end = hasta ?? Today;
        var start = desde ?? Today.AddMonths(-1);
        if (!ModelState.IsValid || clienteId <= 0 || start > end || buscar?.Length > 150) return BadRequest();
        var data = await documents.ReadAsync(Actor, true, null, start, end, "01", buscar?.Trim(), 1, clienteId, null, ct,
            invoiceLookup: true);
        return Json(data.GetProperty("FacturasNotaCredito"));
    }

    [HttpGet("TasaIgv")]
    public async Task<IActionResult> TaxRate(string scope, DateOnly fecha, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (!ModelState.IsValid || fecha == default) return BadRequest();
        var data = await documents.ReadAsync(Actor, scope == "ventas", null, fecha, null, null, null, 1, null, null, ct);
        return Json(new { tasa = CommercialForm.ReadTaxRate(data) });
    }

    [HttpGet("OrdenesCompra")]
    public async Task<IActionResult> PurchaseOrders(string scope, long clienteId, string? buscar, CancellationToken ct)
    {
        if (scope != "ventas" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (!ModelState.IsValid || clienteId <= 0 || buscar?.Length > 150) return BadRequest();
        var data = await documents.ReadAsync(Actor, true, null, null, null, null, buscar?.Trim(), 1, clienteId, null, ct, orderLookup: true);
        return Json(data.GetProperty("OpcionesOrdenCompra"));
    }

    [HttpPost("{id:long}/OrdenCompra")]
    public async Task<IActionResult> LinkPurchaseOrder(string scope, long id, long? ordenCompraId, CancellationToken ct)
    {
        if (scope != "ventas" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (!ModelState.IsValid || ordenCompraId is <= 0) return BadRequest();
        try
        {
            await documents.LinkOrderAsync(Actor, id, ordenCompraId, ct);
            return Notice(scope, id, "Orden de compra actualizada.");
        }
        catch (MaintenanceException error)
        {
            return Notice(scope, id, error.Message);
        }
    }

    [HttpPost("DetraccionMateriales")]
    [RequestSizeLimit(256 * 1024)]
    public async Task<IActionResult> MaterialWithholding(string scope, [FromBody] string[]? codes, CancellationToken ct)
    {
        if (scope != "ventas" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (!ModelState.IsValid || codes is null || codes.Length > 500 ||
            codes.Any(code => string.IsNullOrWhiteSpace(code) || code.Length > 100)) return BadRequest();

        var data = await documents.ReadAsync(Actor, true, null, null, null, null, null, 1, null, null, ct,
            materialCodes: codes);
        return Json(data.TryGetProperty("DetraccionMaterial", out var result) ? (JsonElement?)result : null);
    }

    private const int MaximumBatchBytes = 24 * 1024 * 1024;

    [HttpPost("Importar")]
    [RequestSizeLimit(MaximumBatchBytes + 65536)]
    public async Task<IActionResult> Import(string scope, List<IFormFile> xml, long? id, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (xml.Count is < 1 or > 25 || xml.Sum(file => file.Length) > MaximumBatchBytes || id.HasValue && xml.Count != 1)
            return Notice(scope, id, "Selecciona de 1 a 25 XML, hasta 24 MB en total. Para completar un borrador, selecciona un solo XML.");

        var sales = scope == "ventas";
        var files = new List<(string File, CommercialDocument Document)>(xml.Count);
        foreach (var file in xml)
        {
            var name = Path.GetFileName(file.FileName);
            if (file.Length is <= 0 or > CommercialDocumentXml.MaximumSize)
                return Notice(scope, id, $"{name}: el archivo debe tener contenido y pesar hasta 2 MB.");
            try
            {
                await using var stream = file.OpenReadStream();
                var document = await CommercialDocumentXml.ReadAsync(stream, sales, ct);
                Validate(document);
                files.Add((name, document));
            }
            catch (XmlException) { return Notice(scope, id, $"{name}: no se pudo leer como XML UBL."); }
            catch (FormatException error) { return Notice(scope, id, $"{name}: {error.Message}"); }
        }

        try
        {
            if (files.Count == 1)
            {
                var saved = await documents.SaveAsync(Actor, sales, files[0].Document, id, ct);
                return RedirectToAction(nameof(Details), new { scope, id = saved });
            }
            await documents.ImportBatchAsync(Actor, sales, files, ct);
            return Notice(scope, null, $"Se importaron {files.Count} comprobantes.");
        }
        catch (MaintenanceException error)
        {
            var detail = files.Count == 1 ? $"{files[0].File}: {error.Message}" : error.Message;
            return Notice(scope, id, $"Importación cancelada. {detail}");
        }
    }

    [HttpPost("Guardar")]
    public async Task<IActionResult> Save(string scope, CommercialInput input, long? clienteId, long? presupuestoId, CancellationToken ct)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        var sales = scope == "ventas";
        var data = await documents.ReadAsync(Actor, sales, null, input.Fecha, null, null, null, 1, clienteId, presupuestoId, ct);
        try
        {
            if (!ModelState.IsValid || input.Fecha == default) throw new FormatException("Revisa los campos del comprobante.");
            var rate = CommercialForm.ReadTaxRate(data) ?? throw new FormatException("No hay una tasa de IGV configurada para la fecha de emisión.");
            var document = input.ToDocument(rate, sales);
            if (sales && (!data.TryGetProperty("Cliente", out var client) ||
                CommercialPage.Text(client, "DocumentoTercero") != document.DocumentoTercero))
                throw new FormatException("Selecciona un cliente válido.");
            Validate(document);
            var id = await documents.SaveAsync(Actor, sales, document, null, ct);
            return RedirectToAction(nameof(Details), new { scope, id });
        }
        catch (Exception error) when (error is FormatException or MaintenanceException or OverflowException)
        {
            ModelState.AddModelError("", error is OverflowException ? "Revisa las cantidades y los precios del detalle." : error.Message);
            return View("New", new CommercialForm(scope, data, input, clienteId, presupuestoId, Today));
        }
    }

    [HttpPost("{id:long}/Vincular")]
    public async Task<IActionResult> Link(string scope, long id, long[] ids, long? presupuestoId,
        string? descripcion, decimal? importe, Guid? operacionId, CancellationToken ct, bool popup = false, string? retorno = null)
    {
        if (!Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (scope == "compras" && !Has("COSTOS_PROVEEDOR.EDITAR")) return Forbid();
        if (!ModelState.IsValid || ids.Length > 100 || descripcion?.Length > 300)
            return popup ? BadRequest("Revisa los datos del vínculo.") : Notice(scope, id, "Revisa los datos del vínculo.");
        try
        {
            await documents.LinkAsync(Actor, scope == "ventas", id, ids, presupuestoId, descripcion?.Trim(), importe, ct, operacionId);
            if (popup)
            {
                TempData["CommercialNotice"] = "Vínculos registrados.";
                return Json(new { redirect = LocalReturn(scope, retorno) });
            }
            return Notice(scope, id, "Vínculo registrado.");
        }
        catch (MaintenanceException error)
        {
            return popup ? BadRequest(error.Message) : Notice(scope, id, error.Message);
        }
    }

    [HttpPost("{id:long}/Pago")]
    public async Task<IActionResult> Pay(string scope, long id, bool detraccion, DateOnly fecha, decimal importe,
        string referencia, CancellationToken ct)
    {
        if (scope != "compras" || !Allowed(scope, "CONSULTAR") || !Allowed(scope, "EDITAR")) return Forbid();
        if (!ModelState.IsValid || fecha == default || fecha > Today || string.IsNullOrWhiteSpace(referencia) || referencia.Length > 100)
            return Notice(scope, id, "Revisa la fecha y la referencia del pago.");
        try
        {
            await documents.PayAsync(Actor, id, detraccion, fecha, importe, referencia.Trim(), ct);
            return Notice(scope, id, "Pago registrado.");
        }
        catch (MaintenanceException error) { return Notice(scope, id, error.Message); }
    }

    private IActionResult Notice(string scope, long? id, string message)
    {
        TempData["CommercialNotice"] = message;
        return id.HasValue ? RedirectToAction(nameof(Details), new { scope, id }) : RedirectToAction(nameof(Index), new { scope });
    }

    private static void Validate(CommercialDocument document)
    {
        if (document.Fecha == default || document.Tipo is not ("01" or "03" or "07") ||
            document.Moneda is not ("PEN" or "USD") || document.MonedaDetraccion is not ("PEN" or "USD") ||
            document.Serie?.Length > 30 || document.Numero?.Length > 50 || document.DocumentoTercero.Length is 0 or > 30 ||
            document.NombreTercero.Length > 180 || document.Importe is < 0 or >= 10000000000000000m ||
            document.ImporteDetraccion is < 0 or >= 10000000000000000m || document.CodigoDetraccion?.Length > 30 ||
            document.ReferenciaNota?.Length > 100 || document.MotivoNota?.Length > 300 ||
            document.Detalle.Count is 0 or > 500 || document.Presupuestos.Count > 100 ||
            document.Detalle.Any(line => string.IsNullOrWhiteSpace(line.Descripcion) || line.Descripcion.Length > 1500 ||
                string.IsNullOrWhiteSpace(line.Unidad) || line.Unidad.Length > 30 || line.Codigo?.Length > 100 ||
                line.Cantidad < 0 || line.Precio < 0 || line.Importe < 0))
            throw new FormatException("Revisa la moneda, la identificación, los importes y el detalle del comprobante.");
    }
}
