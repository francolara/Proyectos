using System.Text.Json;
using GestionServicios.Application.Budgets;
using GestionServicios.Web.ViewModels.Shared;
namespace GestionServicios.Web.ViewModels.Budgets;

public sealed class BudgetInputModel
{
    public long? Id { get; set; }
    public long? RevisionId { get; set; }
    public string? HeaderVersion { get; set; }
    public string? RevisionVersion { get; set; }
    public bool NewRevision { get; set; }
    public string? ReturnUrl { get; set; }
    public BudgetDraft Draft { get; set; }=new();
    public BudgetVersion? Original()
    {
        if(!Id.HasValue&&!RevisionId.HasValue)return null;
        var header=Convert.FromBase64String(HeaderVersion??"");var revision=Convert.FromBase64String(RevisionVersion??"");
        if(Id is not >0||RevisionId is not >0||header.Length!=8||revision.Length!=8)throw new FormatException("Falta la versión del registro.");
        return new(Id.Value,RevisionId.Value,"",0,header,revision);
    }
}
public sealed record BudgetListModel(GridModel Grid,FilterModel Filter,bool CanEdit,IReadOnlyList<StateKpi>? Kpis=null);
public sealed record BudgetFormModel(BudgetInputModel Input,BudgetRecord? Record,JsonElement Defaults,bool ReadOnly,bool CanEdit,string ClientLabel,string StoreLabel,int AuditPage=1,bool CanChangeState=false);
public sealed record BudgetLineModel(BudgetDraftLine Line,string Index,string? UnitPrice=null,string? Total=null);
public sealed class BudgetStateInput
{
    public long Id { get; set; }
    public long RevisionId { get; set; }
    public string? HeaderVersion { get; set; }
    public string? RevisionVersion { get; set; }
    public string NewState { get; set; }="";
    public string? EventDate { get; set; }
    public string? Responsible { get; set; }
    public string? Reason { get; set; }
    public string? ReturnUrl { get; set; }
}
