using GestionServicios.Contracts;
using GestionServicios.Web.ViewModels.Shared;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

// Solo se registra en WebApplicationFactory. El producto no referencia este ejecutable.
[AllowAnonymous]
public sealed class UiFixtureController : Controller
{
    [HttpGet("/__tests/maintenance-state")]
    public IActionResult MaintenanceState([FromServices] GestionServicios.Application.Maintenance.IMaintenanceRepository repository)
    {
        var fake=(TestMaintenanceRepository)repository;return Json(new{saved=fake.SaveCount,payload=fake.LastPayload,copies=fake.Copies});
    }
    [HttpGet("/__tests/parts/{part}")]
    public IActionResult Index(string part)
    {
        return part switch
        {
            "date" => PartialView("~/Views/Shared/_DatePicker.cshtml", new DatePickerModel("Fecha", "Fecha de guía", new(2026, 11, 3))),
            "filters" => PartialView("~/Views/Shared/_Filters.cshtml", FilterModel.DocumentDefaults("Home")),
            "trace" => PartialView("~/Views/Shared/_Trace.cshtml", new RecordTrace(new(2026, 11, 3, 18, 30, 0, TimeSpan.Zero), "Operador", null, null, [])),
            "grid" => PartialView("~/Views/Shared/_Grid.cshtml", new GridModel
            {
                Headers = ["Descripción"], Rows = [new("1", ["Registro <prueba>"])], TotalRecords = 21,
                Query = new() { ["Search"] = "buscar" }
            }),
            "view" => PartialView("~/Views/Shared/_RecordForm.cshtml", new RecordFormModel
            {
                ReadOnly = true, FieldsPartial = "_DatePicker", Fields = new DatePickerModel("Fecha", "Fecha", new(2026, 11, 3))
            }),
            "modal" => PartialView("~/Views/Shared/_Modal.cshtml", new ModalModel("test-modal", "Ayuda de material")),
            _ => NotFound()
        };
    }
}
