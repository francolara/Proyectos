using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Globalization;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReportesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, string? preset = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = await ConstruirReporteAsync(baseVm, fechaDesde, fechaHasta, sedeId, preset, incluirDetalle: false);
        return View(vm);
    }

    [HttpGet]
    public async Task<IActionResult> Imprimir(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, string? preset = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var vm = await ConstruirReporteAsync(baseVm, fechaDesde, fechaHasta, sedeId, preset, incluirDetalle: true);
        return View(vm);
    }

    private async Task<ReportesIndexViewModel> ConstruirReporteAsync(
        ModuloBaseViewModel baseVm,
        DateOnly? fechaDesde,
        DateOnly? fechaHasta,
        int? sedeId,
        string? preset,
        bool incluirDetalle)
    {
        var negocioId = baseVm.NegocioId;
        var sedes = await spService.EspaciosComboSedesAsync(negocioId, baseVm.EsAdministrador ? null : baseVm.SedeIdAsignada);
        var sedeFiltro = AplicarSedeAsignada(baseVm, sedeId);
        if (baseVm.EsAdministrador && sedeFiltro.HasValue && !sedes.Any(x => x.Value == sedeFiltro.Value.ToString()))
            sedeFiltro = null;

        var (desde, hasta, presetNormalizado) = ResolverRango(preset, fechaDesde, fechaHasta);
        var diasPeriodo = Math.Max(1, (hasta.DayNumber - desde.DayNumber) + 1);
        var hastaAnterior = desde.AddDays(-1);
        var desdeAnterior = hastaAnterior.AddDays(-(diasPeriodo - 1));

        var ocupacionTask = spService.ReportesOcupacionPorEspacioAsync(negocioId, desde, hasta, sedeFiltro);
        var ingresosTask = spService.ReportesIngresosPorDiaAsync(negocioId, desde, hasta, sedeFiltro);
        var reservasTask = spService.ReportesReservasPorDiaAsync(negocioId, desde, hasta, sedeFiltro);
        var resumenActualTask = spService.ReportesResumenOperativoAsync(negocioId, desde, hasta, sedeFiltro);
        var cobranzaActualTask = spService.ReportesResumenCobranzaAsync(negocioId, desde, hasta, sedeFiltro);

        var ingresosAnteriorTask = spService.ReportesIngresosPorDiaAsync(negocioId, desdeAnterior, hastaAnterior, sedeFiltro);
        var reservasAnteriorTask = spService.ReportesReservasPorDiaAsync(negocioId, desdeAnterior, hastaAnterior, sedeFiltro);
        var resumenAnteriorTask = spService.ReportesResumenOperativoAsync(negocioId, desdeAnterior, hastaAnterior, sedeFiltro);
        var cobranzaAnteriorTask = spService.ReportesResumenCobranzaAsync(negocioId, desdeAnterior, hastaAnterior, sedeFiltro);
        var detallePagosTask = incluirDetalle
            ? spService.ReportesDetallePagosAsync(negocioId, desde, hasta, sedeFiltro)
            : Task.FromResult(new List<ReportePagoDetalleItemViewModel>());
        var detalleReservasTask = incluirDetalle
            ? spService.ReportesDetalleReservasAsync(negocioId, desde, hasta, sedeFiltro)
            : Task.FromResult(new List<ReporteReservaDetalleItemViewModel>());

        await Task.WhenAll(
            ocupacionTask,
            ingresosTask,
            reservasTask,
            resumenActualTask,
            cobranzaActualTask,
            ingresosAnteriorTask,
            reservasAnteriorTask,
            resumenAnteriorTask,
            cobranzaAnteriorTask,
            detallePagosTask,
            detalleReservasTask);

        var vm = new ReportesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            SedeIdAsignada = baseVm.SedeIdAsignada,
            EsAdministrador = baseVm.EsAdministrador,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Preset = presetNormalizado,
            FechaDesde = desde,
            FechaHasta = hasta,
            FechaDesdeAnterior = desdeAnterior,
            FechaHastaAnterior = hastaAnterior,
            DiasPeriodo = diasPeriodo,
            SedeId = sedeFiltro,
            SedesFiltro = PrepararSedesFiltro(sedes, baseVm.EsAdministrador, sedeFiltro),
            Ocupacion = ocupacionTask.Result,
            ReservasPorDia = reservasTask.Result,
            ReservasPeriodoAnterior = reservasAnteriorTask.Result,
            Ingresos = ingresosTask.Result,
            IngresosPeriodoAnterior = ingresosAnteriorTask.Result,
            ResumenActual = resumenActualTask.Result,
            ResumenAnterior = resumenAnteriorTask.Result,
            CobranzaActual = cobranzaActualTask.Result,
            CobranzaAnterior = cobranzaAnteriorTask.Result,
            DetallePagos = detallePagosTask.Result,
            DetalleReservas = detalleReservasTask.Result
        };

        return vm;
    }

    [HttpGet]
    public async Task<IActionResult> ExportCsv(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, string? preset = null, string? bloque = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return Forbid();

        var sedes = await spService.EspaciosComboSedesAsync(negocioId, baseVm.EsAdministrador ? null : baseVm.SedeIdAsignada);
        var sedeFiltro = AplicarSedeAsignada(baseVm, sedeId);
        if (baseVm.EsAdministrador && sedeFiltro.HasValue && !sedes.Any(x => x.Value == sedeFiltro.Value.ToString()))
            sedeFiltro = null;

        var (desde, hasta, _) = ResolverRango(preset, fechaDesde, fechaHasta);
        var bloqueNormalizado = (bloque ?? "todo").Trim().ToLowerInvariant();
        if (bloqueNormalizado is not ("todo" or "resumen" or "ocupacion" or "ingresos"))
            bloqueNormalizado = "todo";

        var resumen = await spService.ReportesResumenOperativoAsync(negocioId, desde, hasta, sedeFiltro);
        var cobranza = await spService.ReportesResumenCobranzaAsync(negocioId, desde, hasta, sedeFiltro);
        var ocupacion = await spService.ReportesOcupacionPorEspacioAsync(negocioId, desde, hasta, sedeFiltro);
        var ingresos = await spService.ReportesIngresosPorDiaAsync(negocioId, desde, hasta, sedeFiltro);
        var reservas = await spService.ReportesReservasPorDiaAsync(negocioId, desde, hasta, sedeFiltro);

        var sb = new StringBuilder();
        const string sep = ";";
        var cultura = new CultureInfo("es-PE");
        var diasPeriodo = Math.Max(1, (hasta.DayNumber - desde.DayNumber) + 1);

        if (bloqueNormalizado is "todo" or "resumen")
        {
            var ticketPromedioCobranza = cobranza.ReservasCobradas > 0 ? cobranza.MontoCobrado / cobranza.ReservasCobradas : 0m;
            var cobranzaPct = resumen.MontoReservado > 0 ? (cobranza.MontoCobrado / resumen.MontoReservado) * 100m : 0m;
            sb.AppendLine("[RESUMEN_OPERATIVO]");
            sb.AppendLine(string.Join(sep, new[]
            {
                "NegocioId","SedeId","FechaDesde","FechaHasta","Dias","TotalReservas","Pendientes","Confirmadas","Pagadas","Canceladas","NoShow",
                "MontoReservado","SaldoPendiente"
            }));
            sb.AppendLine(string.Join(sep, new[]
            {
                negocioId.ToString(),
                sedeFiltro?.ToString() ?? string.Empty,
                desde.ToString("yyyy-MM-dd"),
                hasta.ToString("yyyy-MM-dd"),
                diasPeriodo.ToString(),
                resumen.TotalReservas.ToString(),
                resumen.TotalPendientes.ToString(),
                resumen.TotalConfirmadas.ToString(),
                resumen.TotalPagadas.ToString(),
                resumen.TotalCanceladas.ToString(),
                resumen.TotalNoShow.ToString(),
                FormatoNumero(resumen.MontoReservado, cultura),
                FormatoNumero(resumen.SaldoPendiente, cultura)
            }));
            sb.AppendLine();
            sb.AppendLine("[RESUMEN_COBRANZA]");
            sb.AppendLine(string.Join(sep, new[]
            {
                "NegocioId","SedeId","FechaDesde","FechaHasta","Dias","CantidadPagos","ReservasCobradas","MontoCobrado","TicketPromedioCobranza","CobranzaPctSobreReservado"
            }));
            sb.AppendLine(string.Join(sep, new[]
            {
                negocioId.ToString(),
                sedeFiltro?.ToString() ?? string.Empty,
                desde.ToString("yyyy-MM-dd"),
                hasta.ToString("yyyy-MM-dd"),
                diasPeriodo.ToString(),
                cobranza.CantidadPagos.ToString(),
                cobranza.ReservasCobradas.ToString(),
                FormatoNumero(cobranza.MontoCobrado, cultura),
                FormatoNumero(ticketPromedioCobranza, cultura),
                FormatoNumero(cobranzaPct, cultura)
            }));
            sb.AppendLine();
        }

        if (bloqueNormalizado is "todo" or "ocupacion")
        {
            sb.AppendLine("[OCUPACION]");
            sb.AppendLine(string.Join(sep, new[]
            {
                "SedeId","EspacioDeportivoId","Sede","Espacio","CantidadReservas","HorasReservadas","MontoReservado","MontoCobrado","TicketPromedio","CobranzaPct"
            }));
            foreach (var o in ocupacion)
            {
                var ticket = o.CantidadReservas > 0 ? o.MontoCobrado / o.CantidadReservas : 0m;
                var cobranzaOcupacion = o.MontoReservado > 0 ? (o.MontoCobrado / o.MontoReservado) * 100m : 0m;
                sb.AppendLine(string.Join(sep, new[]
                {
                    o.SedeId.ToString(),
                    o.EspacioDeportivoId.ToString(),
                    EscapeCsv(o.Sede, sep),
                    EscapeCsv(o.Espacio, sep),
                    o.CantidadReservas.ToString(),
                    FormatoNumero(o.HorasReservadas, cultura),
                    FormatoNumero(o.MontoReservado, cultura),
                    FormatoNumero(o.MontoCobrado, cultura),
                    FormatoNumero(ticket, cultura),
                    FormatoNumero(cobranzaOcupacion, cultura)
                }));
            }
            var totalReservasOcup = ocupacion.Sum(x => x.CantidadReservas);
            var totalReservadoOcup = ocupacion.Sum(x => x.MontoReservado);
            var totalCobradoOcup = ocupacion.Sum(x => x.MontoCobrado);
            var ticketTotalOcup = totalReservasOcup > 0 ? totalCobradoOcup / totalReservasOcup : 0m;
            var cobranzaTotalOcup = totalReservadoOcup > 0 ? (totalCobradoOcup / totalReservadoOcup) * 100m : 0m;
            sb.AppendLine(string.Join(sep, new[]
            {
                "TOTAL","","","",
                totalReservasOcup.ToString(),
                FormatoNumero(ocupacion.Sum(x => x.HorasReservadas), cultura),
                FormatoNumero(totalReservadoOcup, cultura),
                FormatoNumero(totalCobradoOcup, cultura),
                FormatoNumero(ticketTotalOcup, cultura),
                FormatoNumero(cobranzaTotalOcup, cultura)
            }));
            sb.AppendLine();
        }

        if (bloqueNormalizado is "todo" or "ingresos")
        {
            sb.AppendLine("[INGRESOS]");
            sb.AppendLine(string.Join(sep, new[] { "FechaPago", "ReservasCobradas", "Ingresos", "TicketPromedioDia" }));
            foreach (var i in ingresos)
            {
                var ticketDia = i.CantidadReservas > 0 ? i.Ingresos / i.CantidadReservas : 0m;
                sb.AppendLine(string.Join(sep, new[]
                {
                    i.Fecha.ToString("yyyy-MM-dd"),
                    i.CantidadReservas.ToString(),
                    FormatoNumero(i.Ingresos, cultura),
                    FormatoNumero(ticketDia, cultura)
                }));
            }
            var totalReservasIngreso = ingresos.Sum(x => x.CantidadReservas);
            var totalIngresos = ingresos.Sum(x => x.Ingresos);
            var ticketPromedioIngreso = totalReservasIngreso > 0 ? totalIngresos / totalReservasIngreso : 0m;
            sb.AppendLine(string.Join(sep, new[]
            {
                "TOTAL",
                totalReservasIngreso.ToString(),
                FormatoNumero(totalIngresos, cultura),
                FormatoNumero(ticketPromedioIngreso, cultura)
            }));
        }

        if (bloqueNormalizado is "todo" or "ingresos")
        {
            sb.AppendLine();
            sb.AppendLine("[RESERVAS_POR_DIA]");
            sb.AppendLine(string.Join(sep, new[] { "FechaReserva", "CantidadReservas", "MontoReservado" }));
            foreach (var r in reservas)
            {
                sb.AppendLine(string.Join(sep, new[]
                {
                    r.Fecha.ToString("yyyy-MM-dd"),
                    r.CantidadReservas.ToString(),
                    FormatoNumero(r.MontoReservado, cultura)
                }));
            }
        }

        var bytes = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        var nombreSede = sedeFiltro.HasValue ? $"_S{sedeFiltro.Value}" : "_ALL";
        var fileName = $"Reporte_{bloqueNormalizado}_{negocioId}{nombreSede}_{desde:yyyyMMdd}_{hasta:yyyyMMdd}.csv";
        return File(bytes, "text/csv; charset=utf-8", fileName);
    }

    private static (DateOnly Desde, DateOnly Hasta, string Preset) ResolverRango(string? preset, DateOnly? fechaDesde, DateOnly? fechaHasta)
    {
        var presetNormalizado = (preset ?? string.Empty).Trim().ToLowerInvariant();
        var hoy = DateOnly.FromDateTime(DateTime.Today);

        if (presetNormalizado == "hoy")
            return (hoy, hoy, "hoy");

        if (presetNormalizado == "30d")
            return (hoy.AddDays(-29), hoy, "30d");

        if (presetNormalizado is "mes" or "mesactual")
        {
            var inicioMes = new DateOnly(hoy.Year, hoy.Month, 1);
            return (inicioMes, hoy, "mes");
        }

        if (fechaDesde.HasValue || fechaHasta.HasValue)
        {
            var desdeManual = fechaDesde ?? hoy.AddDays(-6);
            var hastaManual = fechaHasta ?? hoy;
            if (hastaManual < desdeManual) hastaManual = desdeManual;
            return (desdeManual, hastaManual, "custom");
        }

        return (hoy.AddDays(-6), hoy, "7d");
    }

    private static List<SelectListItem> PrepararSedesFiltro(IEnumerable<SelectListItem> sedes, bool esAdministrador, int? sedeSeleccionada)
    {
        var items = sedes.Select(x => new SelectListItem
        {
            Text = x.Text,
            Value = x.Value,
            Selected = sedeSeleccionada.HasValue && x.Value == sedeSeleccionada.Value.ToString()
        }).ToList();

        if (esAdministrador)
        {
            items.Insert(0, new SelectListItem
            {
                Text = "Todas las sedes",
                Value = string.Empty,
                Selected = !sedeSeleccionada.HasValue
            });
        }

        return items;
    }

    private static string EscapeCsv(string value, string separador)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        if (!value.Contains('"') && !value.Contains(separador) && !value.Contains('\n') && !value.Contains('\r'))
            return value;
        return $"\"{value.Replace("\"", "\"\"")}\"";
    }

    private static string FormatoNumero(decimal valor, CultureInfo cultura)
        => valor.ToString("0.00", cultura);
}
