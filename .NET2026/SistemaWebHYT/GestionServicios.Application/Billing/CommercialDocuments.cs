using System.Text.Json;

namespace GestionServicios.Application.Billing;

public sealed record CommercialLine(string? Codigo, string Descripcion, string Unidad, decimal Cantidad,
    decimal Precio, decimal Importe);

public sealed record CommercialDocument(string Tipo, string? Serie, string? Numero, DateOnly Fecha, string Moneda,
    string DocumentoTercero, string NombreTercero, string FormaPago, decimal Importe, IReadOnlyList<CommercialLine> Detalle,
    decimal ImporteDetraccion, string MonedaDetraccion, string? CodigoDetraccion, string? ReferenciaNota, string? MotivoNota,
    IReadOnlyList<long> Presupuestos, string? RucHYT = null, string? XmlOriginal = null, Guid? OperacionId = null, decimal? DetraccionDocumento = null, DateOnly? FechaVencimiento = null, decimal? PorcentajeDetraccion = null, long? OrdenCompraId = null, string? CodigoMotivoNota = null);

public interface ICommercialDocumentRepository
{
    Task<JsonElement> ReadAsync(string actor, bool sales, long? id, DateOnly? from, DateOnly? to, string? type,
        string? search, int page, long? clientId, long? budgetId, CancellationToken ct, bool withoutFactoring = false, bool? hasFactoring = null,
        IReadOnlyList<string>? materialCodes = null, bool orderLookup = false, bool invoiceLookup = false);
    Task<long> SaveAsync(string actor, bool sales, CommercialDocument document, long? id, CancellationToken ct);
    Task ImportBatchAsync(string actor, bool sales, IReadOnlyList<(string File, CommercialDocument Document)> files, CancellationToken ct);
    Task LinkAsync(string actor, bool sales, long id, IReadOnlyList<long> ids, long? budgetId,
        string? description, decimal? amount, CancellationToken ct, Guid? operationId = null);
    Task LinkOrderAsync(string actor, long id, long? orderId, CancellationToken ct);
    Task<JsonElement> ReadWithholdingsAsync(string actor, string state, string? supplier, string? search, int page, CancellationToken ct);
    Task PayAsync(string actor, long id, bool withholding, DateOnly date, decimal amount, string reference, CancellationToken ct);
}
