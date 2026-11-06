using System.ComponentModel.DataAnnotations;
using System.Text.Json;
namespace GestionServicios.Web.ViewModels.Billing;
public sealed class InvoiceMarkInput
{
    public long OrderId {get;set;}
    public string OrderVersion {get;set;}="";
    public long? ReferenceId {get;set;}
    public string? ReferenceVersion {get;set;}
    public Guid OperationId {get;set;}
    public DateOnly? Date {get;set;}
    [StringLength(2)]public string? Type {get;set;}="01";
    [StringLength(30)]public string? Series {get;set;}
    [StringLength(50)]public string? Number {get;set;}
    [StringLength(1000)]public string? Reason {get;set;}
    public bool Withdraw {get;set;}
    public string? ReturnUrl {get;set;}
}
public sealed record InvoiceMarkModel(InvoiceMarkInput Input,JsonElement Data,bool CanEdit,bool Popup,int Page);