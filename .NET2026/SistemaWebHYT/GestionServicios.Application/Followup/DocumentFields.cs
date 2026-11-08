using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
namespace GestionServicios.Application.Followup;
public sealed record DocumentField(string Name,string Label,string Kind="Text",bool Required=false,int Maximum=100);
public static class DocumentFields
{
    public static IReadOnlyList<DocumentField> For(DocumentKind kind)=>kind.Code switch
    {
        "OC"=>[new("RevisionId","Revisión del presupuesto","Revision",true),new("NumeroOCCliente","N° OC del cliente",Required:true),new("FechaEmisionOC","Fecha de emisión OC","Date",true),new("FechaRecepcionOC","Fecha de recepción OC","Date",true),new("Importe","Importe de la OC","Money"),new("Observacion","Observación","LongText",Maximum:2000)],
        "GUIA"=>[new("FechaActa","Fecha de la guía","Date",true),new("NumeroActa","N° guía",Required:true),new("Observacion","Observación","LongText",Maximum:2000)],
        "CONFIRMACION"=>[new("FechaConfirmacion","Fecha de confirmación","Date",true),new("NumeroConfirmacion","N° confirmación",Required:true),new("RepresentanteCliente","Representante del cliente",Maximum:150),new("Observacion","Observación","LongText",Maximum:2000)],
        "EJECUCION"=>[new("Responsable","Responsable del trabajo",Maximum:150),new("FechaProgramada","Fecha programada","Date"),new("FechaInicioReal","Fecha de inicio real","Date"),new("FechaTerminoReal","Fecha de término real","Date"),new("Descripcion","Descripción del trabajo realizado","LongText",Maximum:2000)],
        _=>throw new ArgumentException("Tipo documental no válido.")
    };
    public static string PrimaryKey(DocumentKind kind)=>kind.Code switch{"OC"=>"OrdenCompraId","GUIA"=>"GuiaId","CONFIRMACION"=>"ConfirmacionId","EJECUCION"=>"EjecucionId",_=>throw new ArgumentException("Tipo documental no válido.")};
    public static (Dictionary<string,object?> Values,IReadOnlyList<(string Field,string Message)> Issues) Validate(DocumentKind kind,IReadOnlyDictionary<string,string?> fields,string? reason,bool editing)
    {
        var values=new Dictionary<string,object?>();var issues=new List<(string,string)>();
        foreach(var field in For(kind))
        {
            var text=fields.GetValueOrDefault(field.Name)?.Trim();values[field.Name]=string.IsNullOrEmpty(text)?null:text;
            if(string.IsNullOrEmpty(text)){if(field.Required)issues.Add((field.Name,"Completa "+field.Label.ToLowerInvariant()+"."));continue;}
            switch(field.Kind)
            {
                case "Date":
                    if(!DateOnly.TryParseExact(text,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out _))issues.Add((field.Name,"Indica una fecha válida."));
                    break;
                case "Money":
                    if(!Regex.IsMatch(text,@"^\d{1,17}(\.\d{1,2})?$",RegexOptions.CultureInvariant)||!decimal.TryParse(text,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out var money))issues.Add((field.Name,"Indica un importe positivo o cero con hasta dos decimales."));
                    else values[field.Name]=money;
                    break;
                case "Revision":
                    if(!long.TryParse(text,NumberStyles.None,CultureInfo.InvariantCulture,out var id)||id<=0)issues.Add((field.Name,"Selecciona una revisión."));else values[field.Name]=id;
                    break;
                default:if(text.Length>field.Maximum)issues.Add((field.Name,$"Admite hasta {field.Maximum} caracteres."));break;
            }
        }
        reason=reason?.Trim();
        if(editing&&string.IsNullOrEmpty(reason))issues.Add(("MotivoCambio","Indica el motivo de la modificación."));
        if(reason?.Length>1000)issues.Add(("MotivoCambio","El motivo admite hasta 1000 caracteres."));
        values["MotivoCambio"]=reason;
        return(values,issues);
    }
}