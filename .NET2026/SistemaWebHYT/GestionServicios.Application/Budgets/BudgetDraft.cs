using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using GestionServicios.Application.Maintenance;

namespace GestionServicios.Application.Budgets;

public sealed record BudgetField(string Name,string Label,string Kind,int Maximum,bool Required=false);
public sealed class BudgetDraft
{
    public long? ClienteId { get; set; }
    public long? TiendaId { get; set; }
    public string TipoAtencion { get; set; }="INVITACION";
    public string? Descripcion { get; set; }
    public string? FechaSolicitud { get; set; }
    public bool DetalleNeto { get;set; }
    public bool RevalorizarMoneda { get; set; }
    public Dictionary<string,string?> Revision { get; set; }=[];
    public List<BudgetDraftLine> Partidas { get; set; }=[];
}
public sealed class BudgetDraftLine
{
    public long? PartidaId { get; set; }
    public long? PartidaTarifarioId { get; set; }
    public long? MaterialId { get; set; }
    public string OrigenPrecio { get; set; }="MANUAL";
    public string? Grupo { get; set; }
    public string? Posicion { get; set; }
    public string? CodigoInterno { get; set; }
    public string? CodigoMaterialSAP { get; set; }
    public string? CentroSAPFICO { get; set; }
    public string? NumeroNecesidad { get; set; }
    public string? TipoServicio { get; set; }
    public string? Descripcion { get; set; }
    public string? Unidad { get; set; }
    public string Cantidad { get; set; }="1.00";
    public string? PrecioManual { get; set; }
    public string CodigoMonedaOrigen { get; set; }="PEN";
    public string CodigoTipoIGVOrigen { get; set; }="SIN_IGV";
    public string? MotivoPrecioManual { get; set; }
    public string? ZonasTexto { get; set; }
}
public static partial class BudgetDraftRules
{
    public static IReadOnlyList<BudgetField> HeaderFields { get; }=
    [
        new("FechaEmision","Fecha de emisión","Date",10,true),
        new("CodigoMoneda","Moneda","Currency",3,true),new("CodigoTipoIGV","Tratamiento IGV de los precios","Tax",10,true),
        new("FechaEjecucion","Fecha de ejecución","Date",10),new("Trimestre","Trimestre","Text",40),new("VigenciaTexto","Vigencia de la oferta","Text",100),
        new("ZonaTexto","Zona informativa","Text",200),new("CodigoSAPProveedor","Código SAP de HYT como proveedor","Text",60),new("ContratoMarco","Contrato marco","Text",60),
        new("Especialidad","Especialidad","Text",150),new("Coordinador","Coordinador","Text",150),new("Encargado","Encargado","Text",150),
        new("Firmante","Firmante","Text",150),new("CargoFirmante","Cargo del firmante","Text",150),new("Consideraciones","Consideraciones","LongText",20000)
    ];
    [GeneratedRegex(@"^\d{1,13}(\.\d{1,2})?$")]
    private static partial Regex DecimalPattern();
    public static bool Number(string? text,out decimal value)
    {
        value=0;
        return text is not null && DecimalPattern().IsMatch(text) && decimal.TryParse(text,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out value);
    }
    public static IReadOnlyList<ValidationIssue> Validate(BudgetDraft draft)
    {
        var issues=new List<ValidationIssue>();
        if(draft.ClienteId is not >0)issues.Add(new("ClienteId","Selecciona el cliente."));
        if(draft.TiendaId is not >0)issues.Add(new("TiendaId","Selecciona una tienda del cliente."));
        if(draft.TipoAtencion is not ("INVITACION" or "EMERGENCIA"))issues.Add(new("TipoAtencion","Selecciona Invitación o Emergencia."));
        if(string.IsNullOrWhiteSpace(draft.Descripcion)||draft.Descripcion.Length>1500)issues.Add(new("Descripcion","Completa la descripción del trabajo (máximo 1500 caracteres)."));
        if(!string.IsNullOrEmpty(draft.FechaSolicitud)&&!DateOnly.TryParseExact(draft.FechaSolicitud,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out _))issues.Add(new("FechaSolicitud","Revisa la fecha de solicitud."));
        foreach(var field in HeaderFields)
        {
            var value=draft.Revision.GetValueOrDefault(field.Name)?.Trim();
            if(string.IsNullOrEmpty(value)) { if(field.Required)issues.Add(new("Revision["+field.Name+"]","Completa "+field.Label.ToLowerInvariant()+"."));continue; }
            if(value.Length>field.Maximum)issues.Add(new("Revision["+field.Name+"]",$"{field.Label}: máximo {field.Maximum} caracteres."));
            if(field.Kind=="Date"&&!DateOnly.TryParseExact(value,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out _))issues.Add(new("Revision["+field.Name+"]","Selecciona una fecha válida."));
            if(field.Kind=="Currency"&&value is not ("PEN" or "USD")||field.Kind=="Tax"&&value is not ("CON_IGV" or "SIN_IGV"))issues.Add(new("Revision["+field.Name+"]","Selecciona una opción válida."));
        }
        if(draft.Partidas.Count>500)issues.Add(new("Partidas","El presupuesto admite hasta 500 partidas."));
        for(var index=0;index<draft.Partidas.Count;index++)
        {
            var line=draft.Partidas[index];
            if(draft.DetalleNeto&&(line.MaterialId is not >0||line.CodigoTipoIGVOrigen!="SIN_IGV"||draft.Revision.GetValueOrDefault("CodigoTipoIGV")!="SIN_IGV"))issues.Add(new("Partidas",$"Ítem {index+1}: selecciona un material y registra el precio sin IGV."));
            var prefix=$"Partidas[{index}].";
            if(!Number(line.Cantidad,out var quantity)||quantity<=0)issues.Add(new(prefix+"Cantidad",$"Partida {index+1}: cantidad positiva con hasta dos decimales."));
            if(!string.IsNullOrEmpty(line.PrecioManual)&&!Number(line.PrecioManual,out _))issues.Add(new(prefix+"PrecioManual",$"Partida {index+1}: precio no negativo con hasta dos decimales."));
            if(line.OrigenPrecio is not ("MANUAL" or "TARIFARIO"))issues.Add(new(prefix+"OrigenPrecio","Selecciona el origen del precio."));
            if(line.OrigenPrecio=="TARIFARIO"&&line.PartidaTarifarioId is not >0)issues.Add(new(prefix+"PartidaTarifarioId","Selecciona la partida del tarifario."));
            if(line.OrigenPrecio=="MANUAL")
            {
                if(line.CodigoTipoIGVOrigen is not ("SIN_IGV" or "CON_IGV"))issues.Add(new(prefix+"CodigoTipoIGVOrigen","Selecciona el tratamiento IGV del precio manual."));
                if(string.IsNullOrWhiteSpace(line.Descripcion)||string.IsNullOrWhiteSpace(line.Unidad))issues.Add(new(prefix+"Descripcion",$"Partida {index+1}: completa descripción y unidad."));
                if(line.PrecioManual is {Length:>0}&&string.IsNullOrWhiteSpace(line.MotivoPrecioManual))issues.Add(new(prefix+"MotivoPrecioManual",$"Partida {index+1}: indica el motivo del precio manual."));
                if(line.CodigoMonedaOrigen!=draft.Revision.GetValueOrDefault("CodigoMoneda"))issues.Add(new(prefix+"CodigoMonedaOrigen",$"Partida {index+1}: revaloriza el precio en la moneda del presupuesto."));
            }
            foreach(var (name,value,maximum) in new (string,string?,int)[] {("Descripcion",line.Descripcion,1500),("Unidad",line.Unidad,30),("Grupo",line.Grupo,200),("Posicion",line.Posicion,60),("CodigoInterno",line.CodigoInterno,60),("CodigoMaterialSAP",line.CodigoMaterialSAP,60),("CentroSAPFICO",line.CentroSAPFICO,50),("NumeroNecesidad",line.NumeroNecesidad,60),("TipoServicio",line.TipoServicio,150),("MotivoPrecioManual",line.MotivoPrecioManual,300),("ZonasTexto",line.ZonasTexto,1000)})
                if(value?.Length>maximum)issues.Add(new(prefix+name,$"Partida {index+1}, {name}: máximo {maximum} caracteres."));
        }
        return issues;
    }
    public static string Payload(BudgetDraft draft)=>JsonSerializer.Serialize(new { draft.ClienteId,draft.TiendaId,draft.TipoAtencion,draft.Descripcion,draft.FechaSolicitud,draft.RevalorizarMoneda,draft.DetalleNeto,
        Revision=HeaderFields.ToDictionary(field=>field.Name,field=>draft.Revision.GetValueOrDefault(field.Name)?.Trim()),draft.Partidas });
}
