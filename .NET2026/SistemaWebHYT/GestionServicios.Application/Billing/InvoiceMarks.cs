using System.Text.Json;
namespace GestionServicios.Application.Billing;
public sealed record InvoiceMarkCommand(long OrderId,byte[] OrderVersion,Guid OperationId,DateOnly? Date,string? Type,string? Series,string? Number,long? ReferenceId,byte[]? ReferenceVersion,bool Withdraw,string? Reason);
public interface IInvoiceMarkRepository
{
    Task<JsonElement?> GetAsync(string actor,long orderId,int page,CancellationToken ct);
    Task<long> SaveAsync(string actor,InvoiceMarkCommand input,CancellationToken ct);
}