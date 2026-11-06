using System.Data.Common;
using GestionServicios.Web.Controllers;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Controllers;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.AspNetCore.Mvc.ModelBinding;

namespace GestionServicios.Web;

public sealed class MaintenanceUnavailableFilter(ILogger<MaintenanceUnavailableFilter> logger, IModelMetadataProvider metadata) : IExceptionFilter
{
    public void OnException(ExceptionContext context)
    {
        if (context.ActionDescriptor is not ControllerActionDescriptor action ||
            action.ControllerTypeInfo.AsType() != typeof(MasterController) && action.ControllerTypeInfo.AsType() != typeof(TaxRateController) && action.ControllerTypeInfo.AsType() != typeof(BudgetController)
            && action.ControllerTypeInfo.AsType()!=typeof(CompanyController) && action.ControllerTypeInfo.AsType()!=typeof(BudgetReportsController) && action.ControllerTypeInfo.AsType()!=typeof(FollowupController) && action.ControllerTypeInfo.AsType()!=typeof(InvoiceMarkController) && action.ControllerTypeInfo.AsType()!=typeof(ReportsController) && action.ControllerTypeInfo.AsType()!=typeof(HomeController)) return;
        if (context.Exception is not DbException && !(context.Exception is InvalidOperationException && context.Exception.Message.StartsWith("Configure la conexión de HYT", StringComparison.Ordinal))) return;
        logger.LogError(context.Exception, "El mantenimiento no pudo acceder al repositorio");
        const string message = "No se pudo completar la consulta. Si estabas guardando, revisa el listado antes de reintentar para comprobar si el registro se guardó.";
        if (context.HttpContext.Request.GetTypedHeaders().Accept?.Any(header => header.MediaType.Value == "application/json") == true)
            context.Result = new ObjectResult(new ProblemDetails { Status = 503, Title = "Servicio temporalmente no disponible", Detail = message }) { StatusCode = 503 };
        else context.Result = new ViewResult
        {
            ViewName = "~/Views/Shared/MaintenanceUnavailable.cshtml", StatusCode = 503,
            ViewData = new ViewDataDictionary(metadata, context.ModelState) { ["Title"] = "Servicio no disponible", ["Message"] = message }
        };
        context.ExceptionHandled = true;
    }
}
