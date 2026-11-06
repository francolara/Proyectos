using System.Net;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Mvc.Testing;

static class MaintenanceRevisionChecks
{
    static string Input(string html,string name)=>Regex.Match(html,"<input[^>]*name=\""+Regex.Escape(name)+"\"[^>]*>").Value;
    static string Value(string html,string name)=>WebUtility.HtmlDecode(Regex.Match(Input(html,name),"value=\"([^\"]*)\"").Groups[1].Value);
    public static async Task Run(Action<bool,string> check)
    {
        var repository=new TestAccessRepository();repository.Account=repository.Account with{Permissions=["CLIENTES.CONSULTAR","CLIENTES.EDITAR","TIENDAS.CONSULTAR","TIENDAS.EDITAR","UNIDADES.CONSULTAR","UNIDADES.EDITAR","ZONAS.CONSULTAR","ZONAS.EDITAR","MATERIALES.CONSULTAR","MATERIALES.EDITAR","TARIFARIOS.CONSULTAR","TARIFARIOS.EDITAR","PRESUPUESTOS.CONSULTAR","OC_CLIENTE.CONSULTAR","GUIAS.CONSULTAR"]};
        using var factory=new HytFactory(repository);using var client=factory.CreateClient(new WebApplicationFactoryClientOptions{BaseAddress=new Uri("https://localhost"),AllowAutoRedirect=false});
        async Task<string> Token(string path)=>Value(await client.GetStringAsync(path),"__RequestVerificationToken");
        await client.PostAsync("/Acceso",new FormUrlEncodedContent(new Dictionary<string,string>{["UserName"]="operador",["Password"]=repository.Password,["__RequestVerificationToken"]=await Token("/Acceso")}));
        var globalList=await client.GetStringAsync("/Tarifarios?search=MAT");
        check(globalList.Contains("MAT-000001")&&globalList.Contains("Material de prueba")&&globalList.Contains("100.00"),"Tarifarios muestra los mismos precios y material que la ficha");
        check((await client.GetAsync("/Maestros/tarifarios")).StatusCode==HttpStatusCode.Redirect,"Menú anterior redirige al listado unificado");
        var globalForm=await client.GetStringAsync("/Tarifarios/Nuevo");
        check(globalForm.Contains("data-open-lookup")&&globalForm.Contains("selectedMaterialId")&&!globalForm.Contains("VersionId"),"Nuevo tarifario permite elegir material mediante ayuda sin versiones");
        var globalSaved=await client.PostAsync("/Tarifarios/Guardar",new FormUrlEncodedContent(new Dictionary<string,string>{["__RequestVerificationToken"]=Value(globalForm,"__RequestVerificationToken"),["selectedMaterialId"]="1",["TaxCode"]="CON_IGV",["PricePEN"]="24.50",["PriceUSD"]="7.25",["Active"]="true"}));
        check(globalSaved.StatusCode==HttpStatusCode.Redirect&&globalSaved.Headers.Location?.ToString()=="/Tarifarios"&&factory.Maintenance.LastPayload!.Contains("24.50"),"Alta desde menú usa el mismo registro de precio y regresa al listado");
        var html=await client.GetStringAsync("/Maestros/clientes/Nuevo");
        check(html.Contains("<output id=\"internal-record-code\">Se asigna al guardar</output>")&&!html.Contains("Fields[Codigo]")&&html.IndexOf("internal-record-code")<html.IndexOf("class=\"hyt-panel hyt-master-form\""),"Cliente nuevo muestra código interno automático no editable ni obligatorio");
        check(WebUtility.HtmlDecode(html).Contains("Código SAP")&&Input(html,"Fields[CodigoSAP]").Length>0&&!Input(html,"Fields[CodigoSAP]").Contains("readonly"),"Código SAP editable se presenta en el cliente");
        check(html.Contains("data-client-link-add")&&html.Contains("data-client-link-template")&&!html.Contains("Fields[UrlCargaDocumentos]"),"Cliente sustituye enlace único por lista con descripción y apertura");
        var clientDetails=await client.GetStringAsync("/Maestros/clientes/1");
        check(!clientDetails.Contains("/Maestros/proveedor-cliente")&&!clientDetails.Contains("Datos de HYT por cliente"),"Cliente oculta el acceso a datos de HYT por cliente");
        foreach(var path in new[]{"/Maestros/clientes/1","/Maestros/materiales/1","/Materiales/1/Tarifario/1","/Presupuestos/1","/Seguimiento/guias/1","/Seguimiento/oc-cliente/1"})
        {
            var detail=WebUtility.HtmlDecode(await client.GetStringAsync(path));
            check(!detail.Contains("Historial de cambios")&&!detail.Contains("Historial de acciones")&&!detail.Contains("Historial de la revisión y del expediente")&&detail.Contains("Fechas y responsables"),"Consulta conserva responsables y oculta historial: "+path);
        }
        var oldClientSection=await client.GetAsync("/Maestros/proveedor-cliente?parentId=1");
        check(oldClientSection.StatusCode==HttpStatusCode.Redirect&&oldClientSection.Headers.Location!.OriginalString.Contains("/Maestros/clientes/1"),"Acceso anterior redirige a la ficha del cliente");
        foreach(var relatedEntity in new[]{"tiendas","contactos"})
        {
            var relatedList=WebUtility.HtmlDecode(await client.GetStringAsync("/Maestros/"+relatedEntity));
            check(relatedList.Contains("/modulo/"+relatedEntity)&&relatedList.Contains("Código de cliente")&&relatedList.Contains("CLI-000099")&&relatedList.Contains("Cliente de prueba"),"Menú y listado muestran cliente en "+relatedEntity);
        }
        html=await client.GetStringAsync("/Maestros/materiales/Nuevo");
        check(!html.Contains("Fields[Posicion]")&&html.Contains("internal-record-code")&&!html.Contains("Fields[CodigoInterno]")&&!Input(html,"Fields[CodigoMaterialSAP]").Contains("readonly"),"Material retira posición y distingue código interno de SAP editable");
        check(html.Contains("/modulo/especialidades")&&html.Contains("/modulo/tipos-servicio"),"Especialidades y tipos de servicio figuran en el menú principal");
        html=await client.GetStringAsync("/Maestros/materiales");check(!html.Contains("href=\"/Maestros/especialidades\"")&&!html.Contains("href=\"/Maestros/tipos-servicio\""),"Materiales ya no contiene enlaces de mantenimiento anidados");
        html=await client.GetStringAsync("/Maestros/unidades/Nuevo");
        check(html.Contains("Abreviatura")&&html.Contains("internal-record-code")&&!Input(html,"Fields[Codigo]").Contains("readonly"),"Unidad distingue identificador automático de abreviatura editable");
        html=await client.GetStringAsync("/Maestros/materiales/1/Editar");
        check(html.Contains("material-tariffs")&&html.Contains("Agregar tarifario")&&!html.Contains("Ver versión")&&!html.Contains("Crear tarifario y precio"),"Ficha de material integra precios y acciones en la misma ventana");
        html=await client.GetStringAsync("/Materiales/1/Tarifario/Nuevo");
        check(html.Contains("Agregar tarifario")&&Input(html,"PricePEN").Length>0&&Input(html,"PriceUSD").Length>0&&!html.Contains("VersionId")&&!html.Contains("Fields[MaterialId]")&&!html.Contains("Posición")&&!html.Contains("Zonas"),"Agregar tarifario solo muestra IGV, precios y estado");
        var simpleFields=new Dictionary<string,string>{["__RequestVerificationToken"]=Value(html,"__RequestVerificationToken"),["TaxCode"]="CON_IGV",["PricePEN"]="12.34",["PriceUSD"]="",["Active"]="true"};
        var savedSimple=await client.PostAsync("/Materiales/1/Tarifario/Guardar",new FormUrlEncodedContent(simpleFields));
        check(savedSimple.StatusCode==HttpStatusCode.OK&&(await savedSimple.Content.ReadAsStringAsync()).Contains("data-price-frame-saved")&&factory.Maintenance.LastPayload!.Contains("\"PrecioUSD\":null"),"Tarifario simple guarda y refresca ficha conservando USD opcional");
        simpleFields["PricePEN"]="1.234";var priorSimple=factory.Maintenance.SaveCount;
        var invalidSimple=await client.PostAsync("/Materiales/1/Tarifario/Guardar",new FormUrlEncodedContent(simpleFields));
        check(invalidSimple.StatusCode==HttpStatusCode.OK&&factory.Maintenance.SaveCount==priorSimple,"Tarifario simple rechaza más de dos decimales");
        check((await client.GetAsync("/Materiales/2/Tarifario/1/Editar")).StatusCode==HttpStatusCode.NotFound,"Tarifario no permite editar un precio de otro material");
        var protectedSimple=await client.GetStringAsync("/Materiales/1/Tarifario/2/Editar");
        check(!protectedSimple.Contains("Guardar tarifario"),"Publicaciones anteriores conservan su protección");
        var oldPopup=await client.GetAsync("/Maestros/partidas-tarifario/Nuevo?compact=true&materialId=1");
        check(oldPopup.StatusCode==HttpStatusCode.Redirect&&oldPopup.Headers.Location!.OriginalString.Contains("/Materiales/1/Tarifario/Nuevo"),"Ayuda anterior redirige a ventana simple");
        repository.Account=repository.Account with{Permissions=["MATERIALES.CONSULTAR","MATERIALES.EDITAR"]};
        html=await client.GetStringAsync("/Maestros/materiales/1/Editar");
        var forbidden=await client.GetAsync("/Maestros/materiales/1/Tarifas");
        check(!html.Contains("data-material-tariffs")&&(forbidden.StatusCode==HttpStatusCode.Forbidden||forbidden.StatusCode==HttpStatusCode.Redirect),"Ficha y endpoint ocultan precios sin permiso de tarifarios");
    }
}
