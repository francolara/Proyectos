// Firma: FRANCO LARA - 06/11/2026 | Nombres visibles y opciones de estado; compatibilidad con códigos históricos.
using GestionServicios.Application.Budgets;
namespace GestionServicios.Web.ViewModels.Budgets;

public static class BudgetStates
{
    public static string Label(string state)=>state switch
    {
        "BORRADOR"=>"Registrado", "ENVIADO"=>"Enviado", "OBSERVADO"=>"Observado",
        "APROBADO"=>"Aceptada", "RECHAZADO"=>"Rechazada", "CANCELADO"=>"Cancelado",
        "VENCIDO"=>"Vencido (histórico)", _=>state
    };
    public static bool NeedsReason(string state)=>state is "OBSERVADO" or "RECHAZADO" or "CANCELADO";
    public static string[] Next(string state,bool hasOrder=false,bool emergency=false)=>hasOrder
        ? emergency ? state switch { "BORRADOR"=>["ENVIADO"],"ENVIADO"=>["APROBADO"],_=>[] } : []
        : state switch
        {
            "BORRADOR" or "OBSERVADO"=>["ENVIADO","CANCELADO"],
            "ENVIADO"=>["OBSERVADO","APROBADO","RECHAZADO","CANCELADO"],
            "APROBADO" or "RECHAZADO" or "VENCIDO"=>["CANCELADO"], _=>[]
        };
    public static string[] Next(BudgetRecord record)=>Next(BudgetRecord.Text(record.Revision,"Estado"),
        record.Text("OrdenCompraId")!="",record.Text("TipoAtencion")=="EMERGENCIA"&&record.Text("EstadoOC")!="CANCELADA"&&record.Text("RevisionOC")==BudgetRecord.Text(record.Revision,"RevisionId"));
}
public sealed record BudgetStateModel(BudgetStateInput Input,string Number,string CurrentState,IReadOnlyList<string> Options,string? Error=null);
