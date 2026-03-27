using System.Text;
using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReportesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today.AddDays(-6));
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today);
        if (hasta < desde) hasta = desde;

        var vm = new ReportesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            FechaDesde = desde,
            FechaHasta = hasta,
            Ocupacion = await spService.ReportesOcupacionPorEspacioAsync(negocioId, desde, hasta),
            Ingresos = await spService.ReportesIngresosPorDiaAsync(negocioId, desde, hasta)
        };

        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> ExportCsv(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        var desde = fechaDesde ?? DateOnly.FromDateTime(DateTime.Today.AddDays(-6));
        var hasta = fechaHasta ?? DateOnly.FromDateTime(DateTime.Today);
        if (hasta < desde) hasta = desde;

        var ocupacion = await spService.ReportesOcupacionPorEspacioAsync(negocioId, desde, hasta);
        var ingresos = await spService.ReportesIngresosPorDiaAsync(negocioId, desde, hasta);

        var sb = new StringBuilder();
        sb.AppendLine("Tipo,Sede,Espacio,Fecha,CantidadReservas,HorasReservadas,MontoReservado,MontoCobrado,Ingresos");

        foreach (var o in ocupacion)
        {
            sb.AppendLine($"OCUPACION,{EscapeCsv(o.Sede)},{EscapeCsv(o.Espacio)},,{o.CantidadReservas},{o.HorasReservadas:F2},{o.MontoReservado:F2},{o.MontoCobrado:F2},");
        }

        foreach (var i in ingresos)
        {
            sb.AppendLine($"INGRESOS,,,{i.Fecha:yyyy-MM-dd},{i.CantidadReservas},,,,{i.Ingresos:F2}");
        }

        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        var fileName = $"Reporte_{negocioId}_{desde:yyyyMMdd}_{hasta:yyyyMMdd}.csv";
        return File(bytes, "text/csv; charset=utf-8", fileName);
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}
