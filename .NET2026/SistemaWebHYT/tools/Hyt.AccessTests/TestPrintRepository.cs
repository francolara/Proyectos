using System.Security.Cryptography;
using System.Text;
using GestionServicios.Application.Budgets;
sealed class TestPrintRepository:IBudgetPrintRepository
{
    private readonly Dictionary<long,BudgetPrint> documents=[];
    public int Count=>documents.Count;
    public Task<long> SaveAsync(string actor,long revisionId,byte[] expectedVersion,string format,string templateVersion,string html,CancellationToken ct)
    {
        var id=(long)documents.Count+1;documents[id]=new(id,revisionId,revisionId,format,templateVersion,DateTimeOffset.UtcNow,"Operador",html,Convert.ToHexString(SHA256.HashData(Encoding.Unicode.GetBytes(html))));return Task.FromResult(id);
    }
    public Task<BudgetPrint?> GetAsync(string actor,long printId,CancellationToken ct)=>Task.FromResult(documents.GetValueOrDefault(printId));
    public Task<BudgetPrintPage> ListAsync(string actor,long revisionId,int page,CancellationToken ct)
    {
        var rows=documents.Values.Where(row=>row.RevisionId==revisionId).OrderByDescending(row=>row.Id).ToArray();
        return Task.FromResult(new BudgetPrintPage(rows.Skip((page-1)*20).Take(20).Select(row=>new BudgetPrintEntry(row.Id,row.RevisionId,row.BudgetId,row.Format,row.TemplateVersion,row.CreatedAt,row.User)).ToArray(),rows.Length,page));
    }
}
