using GestionServicios.Application.Budgets;
using GestionServicios.Web.ViewModels.Budgets;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
namespace GestionServicios.Web.Services;

public sealed class BudgetPrintRenderer(ICompositeViewEngine engine,ITempDataProvider tempData,IModelMetadataProvider metadata,IWebHostEnvironment environment,TimeProvider clock)
{
    public async Task<string> RenderAsync(ControllerContext context,BudgetRecord record,string format,CancellationToken ct)
    {
        var logo="data:image/png;base64,"+Convert.ToBase64String(await File.ReadAllBytesAsync(Path.Combine(environment.WebRootPath,"images","logo-hyt.png"),ct));
        var css=await File.ReadAllTextAsync(Path.Combine(environment.WebRootPath,"hyt-report.css"),ct);
        var model=BudgetPrintModel.Create(record,format,logo,css,clock.GetUtcNow());
        var view=engine.GetView(null,"/Views/BudgetReports/Print.cshtml",false);
        if(!view.Success)throw new InvalidOperationException("No se encontró la plantilla HTML del presupuesto.");
        using var writer=new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
        var data=new ViewDataDictionary<BudgetPrintModel>(metadata,new ModelStateDictionary()){Model=model};
        var viewContext=new ViewContext(context,view.View,data,new TempDataDictionary(context.HttpContext,tempData),writer,new HtmlHelperOptions());
        await view.View.RenderAsync(viewContext);return writer.ToString().TrimStart();
    }
}
