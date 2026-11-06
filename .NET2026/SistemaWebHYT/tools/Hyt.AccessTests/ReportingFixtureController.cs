using GestionServicios.Application.Reporting;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
// Control de datos de prueba registrado solo por WebApplicationFactory; no forma parte del producto publicado.
[Authorize(Policy="perm:REPORTES.CONSULTAR")]
public sealed class ReportingFixtureController:Controller
{
    [HttpPost("/__tests/reporting")]
    public IActionResult Configure([FromServices] IReportingRepository repository,bool empty=false,bool budgets=true,bool orders=true,int rows=1)
    {
        if(repository is not TestReportingRepository fixture)return NotFound();
        fixture.Empty=empty;fixture.ShowBudgets=budgets;fixture.ShowOrders=orders;fixture.ReportRows=Math.Clamp(rows,0,20);return Ok();
    }
}