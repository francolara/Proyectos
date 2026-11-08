using GestionServicios.Application.Budgets;
using GestionServicios.Application.Maintenance;
using System.Security.Claims;
using GestionServicios.Web.ViewModels.Budgets;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
namespace GestionServicios.Web.Services;

public sealed class BudgetPrintRenderer(ICompositeViewEngine engine,ITempDataProvider tempData,IModelMetadataProvider metadata,IWebHostEnvironment environment,TimeProvider clock,ICompanySettingsRepository companies,IDocumentFileStore files)
{
    public async Task<string> RenderAsync(ControllerContext context,BudgetRecord record,string format,CancellationToken ct)
    {
        var logo="data:image/png;base64,"+Convert.ToBase64String(await File.ReadAllBytesAsync(Path.Combine(environment.WebRootPath,"images","logo-hyt.png"),ct));
        var css=await File.ReadAllTextAsync(Path.Combine(environment.WebRootPath,"hyt-report.css"),ct);
        var model=BudgetPrintModel.Create(record,format,logo,css,clock.GetUtcNow());
        var company=await companies.GetAsync(context.HttpContext.User.FindFirstValue(ClaimTypes.NameIdentifier)!,1,ct);
        model.Subtitle=company?.Text("SubtituloCotizacion")??"";
        var signature=BudgetSignature.Metadata(company);
        if(signature is not null)
        {
            var bytes=await files.ReadAsync(signature,ct);
            if(bytes is null)throw new MaintenanceException("La firma configurada no está disponible en Cloudflare. Revisa Configuración HYT.");
            model.Signature="data:"+signature.TipoContenido+";base64,"+Convert.ToBase64String(bytes);
        }
        var view=engine.GetView(null,"/Views/BudgetReports/Print.cshtml",false);
        if(!view.Success)throw new InvalidOperationException("No se encontró la plantilla HTML del presupuesto.");
        using var writer=new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
        var data=new ViewDataDictionary<BudgetPrintModel>(metadata,new ModelStateDictionary()){Model=model};
        var viewContext=new ViewContext(context,view.View,data,new TempDataDictionary(context.HttpContext,tempData),writer,new HtmlHelperOptions());
        await view.View.RenderAsync(viewContext);return writer.ToString().TrimStart();
    }
}
