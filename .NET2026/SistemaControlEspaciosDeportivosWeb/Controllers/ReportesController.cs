using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Globalization;
using ClosedXML.Excel;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ReportesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, string? preset = null)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "REPORTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

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

        await Task.WhenAll(ocupacionTask, ingresosTask, reservasTask, resumenActualTask, cobranzaActualTask, ingresosAnteriorTask, reservasAnteriorTask, resumenAnteriorTask, cobranzaAnteriorTask);

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
            CobranzaAnterior = cobranzaAnteriorTask.Result
        };

        return View(vm);
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

    [HttpGet]
    public async Task<IActionResult> ExportExcel(int negocioId, DateOnly? fechaDesde, DateOnly? fechaHasta, int? sedeId, string? preset = null, string? bloque = null)
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

        using var wb = new XLWorkbook();
        var headerBg = XLColor.FromHtml("#E8F0FE");
        var headerFg = XLColor.FromHtml("#123A74");

        if (bloqueNormalizado is "todo" or "resumen")
        {
            var ws = wb.Worksheets.Add("Resumen");
            ws.Cell(1, 1).Value = "Negocio";
            ws.Cell(1, 2).Value = baseVm.NegocioNombre;
            ws.Cell(2, 1).Value = "Sede";
            ws.Cell(2, 2).Value = sedeFiltro?.ToString() ?? "Todas";
            ws.Cell(3, 1).Value = "Rango";
            ws.Cell(3, 2).Value = $"{desde:dd/MM/yyyy} - {hasta:dd/MM/yyyy}";

            var headers = new[]
            {
                "TotalReservas","Pendientes","Confirmadas","Pagadas","Canceladas","NoShow","MontoReservado","SaldoPendiente","CantidadPagos","ReservasCobradas","MontoCobrado","TicketPromedioCobranza","CobranzaPct"
            };
            for (var i = 0; i < headers.Length; i++)
                ws.Cell(5, i + 1).Value = headers[i];

            var ticketPromedio = cobranza.ReservasCobradas > 0 ? cobranza.MontoCobrado / cobranza.ReservasCobradas : 0m;
            var cobranzaPct = resumen.MontoReservado > 0 ? (cobranza.MontoCobrado / resumen.MontoReservado) * 100m : 0m;
            ws.Cell(6, 1).Value = resumen.TotalReservas;
            ws.Cell(6, 2).Value = resumen.TotalPendientes;
            ws.Cell(6, 3).Value = resumen.TotalConfirmadas;
            ws.Cell(6, 4).Value = resumen.TotalPagadas;
            ws.Cell(6, 5).Value = resumen.TotalCanceladas;
            ws.Cell(6, 6).Value = resumen.TotalNoShow;
            ws.Cell(6, 7).Value = resumen.MontoReservado;
            ws.Cell(6, 8).Value = resumen.SaldoPendiente;
            ws.Cell(6, 9).Value = cobranza.CantidadPagos;
            ws.Cell(6, 10).Value = cobranza.ReservasCobradas;
            ws.Cell(6, 11).Value = cobranza.MontoCobrado;
            ws.Cell(6, 12).Value = ticketPromedio;
            ws.Cell(6, 13).Value = cobranzaPct / 100m;

            FormatearHeader(ws.Range(5, 1, 5, headers.Length), headerBg, headerFg);
            ws.Range(6, 7, 6, 12).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(6, 13).Style.NumberFormat.Format = "0.00%";
            ws.Columns().AdjustToContents();
        }

        if (bloqueNormalizado is "todo" or "ocupacion")
        {
            var ws = wb.Worksheets.Add("Ocupacion");
            var headers = new[]
            {
                "SedeId","EspacioDeportivoId","Sede","Espacio","CantidadReservas","HorasReservadas","MontoReservado","MontoCobrado","TicketPromedio","CobranzaPct"
            };
            for (var i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];
            FormatearHeader(ws.Range(1, 1, 1, headers.Length), headerBg, headerFg);

            var row = 2;
            foreach (var o in ocupacion)
            {
                var ticket = o.CantidadReservas > 0 ? o.MontoCobrado / o.CantidadReservas : 0m;
                var cobranzaOcupacion = o.MontoReservado > 0 ? (o.MontoCobrado / o.MontoReservado) * 100m : 0m;
                ws.Cell(row, 1).Value = o.SedeId;
                ws.Cell(row, 2).Value = o.EspacioDeportivoId;
                ws.Cell(row, 3).Value = o.Sede;
                ws.Cell(row, 4).Value = o.Espacio;
                ws.Cell(row, 5).Value = o.CantidadReservas;
                ws.Cell(row, 6).Value = o.HorasReservadas;
                ws.Cell(row, 7).Value = o.MontoReservado;
                ws.Cell(row, 8).Value = o.MontoCobrado;
                ws.Cell(row, 9).Value = ticket;
                ws.Cell(row, 10).Value = cobranzaOcupacion / 100m;
                row++;
            }

            if (ocupacion.Count > 0)
            {
                ws.Cell(row, 1).Value = "TOTAL";
                ws.Cell(row, 5).Value = ocupacion.Sum(x => x.CantidadReservas);
                ws.Cell(row, 6).Value = ocupacion.Sum(x => x.HorasReservadas);
                ws.Cell(row, 7).Value = ocupacion.Sum(x => x.MontoReservado);
                ws.Cell(row, 8).Value = ocupacion.Sum(x => x.MontoCobrado);
                ws.Range(row, 1, row, headers.Length).Style.Font.Bold = true;
            }

            ws.Column(6).Style.NumberFormat.Format = "#,##0.00";
            ws.Range(2, 7, Math.Max(2, row), 9).Style.NumberFormat.Format = "#,##0.00";
            ws.Range(2, 10, Math.Max(2, row), 10).Style.NumberFormat.Format = "0.00%";
            ws.Columns().AdjustToContents();
        }

        if (bloqueNormalizado is "todo" or "ingresos")
        {
            var ws = wb.Worksheets.Add("Ingresos");
            var headers = new[] { "Fecha", "CantidadReservas", "Ingresos", "TicketPromedioDia" };
            for (var i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];
            FormatearHeader(ws.Range(1, 1, 1, headers.Length), headerBg, headerFg);

            var row = 2;
            foreach (var i in ingresos)
            {
                var ticketDia = i.CantidadReservas > 0 ? i.Ingresos / i.CantidadReservas : 0m;
                ws.Cell(row, 1).Value = i.Fecha.ToDateTime(TimeOnly.MinValue);
                ws.Cell(row, 2).Value = i.CantidadReservas;
                ws.Cell(row, 3).Value = i.Ingresos;
                ws.Cell(row, 4).Value = ticketDia;
                row++;
            }

            if (ingresos.Count > 0)
            {
                ws.Cell(row, 1).Value = "TOTAL";
                ws.Cell(row, 2).Value = ingresos.Sum(x => x.CantidadReservas);
                ws.Cell(row, 3).Value = ingresos.Sum(x => x.Ingresos);
                var totalReservas = ingresos.Sum(x => x.CantidadReservas);
                var totalIngresos = ingresos.Sum(x => x.Ingresos);
                ws.Cell(row, 4).Value = totalReservas > 0 ? totalIngresos / totalReservas : 0m;
                ws.Range(row, 1, row, headers.Length).Style.Font.Bold = true;
            }

            ws.Column(1).Style.DateFormat.Format = "dd/MM/yyyy";
            ws.Range(2, 3, Math.Max(2, row), 4).Style.NumberFormat.Format = "#,##0.00";
            ws.Columns().AdjustToContents();

            var wsReservas = wb.Worksheets.Add("Reservas");
            var headersReservas = new[] { "FechaReserva", "CantidadReservas", "MontoReservado" };
            for (var i = 0; i < headersReservas.Length; i++)
                wsReservas.Cell(1, i + 1).Value = headersReservas[i];
            FormatearHeader(wsReservas.Range(1, 1, 1, headersReservas.Length), headerBg, headerFg);

            var rowReserva = 2;
            foreach (var r in reservas)
            {
                wsReservas.Cell(rowReserva, 1).Value = r.Fecha.ToDateTime(TimeOnly.MinValue);
                wsReservas.Cell(rowReserva, 2).Value = r.CantidadReservas;
                wsReservas.Cell(rowReserva, 3).Value = r.MontoReservado;
                rowReserva++;
            }

            if (reservas.Count > 0)
            {
                wsReservas.Cell(rowReserva, 1).Value = "TOTAL";
                wsReservas.Cell(rowReserva, 2).Value = reservas.Sum(x => x.CantidadReservas);
                wsReservas.Cell(rowReserva, 3).Value = reservas.Sum(x => x.MontoReservado);
                wsReservas.Range(rowReserva, 1, rowReserva, headersReservas.Length).Style.Font.Bold = true;
            }

            wsReservas.Column(1).Style.DateFormat.Format = "dd/MM/yyyy";
            wsReservas.Range(2, 3, Math.Max(2, rowReserva), 3).Style.NumberFormat.Format = "#,##0.00";
            wsReservas.Columns().AdjustToContents();
        }

        await using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;
        var nombreSede = sedeFiltro.HasValue ? $"_S{sedeFiltro.Value}" : "_ALL";
        var fileName = $"Reporte_{bloqueNormalizado}_{negocioId}{nombreSede}_{desde:yyyyMMdd}_{hasta:yyyyMMdd}.xlsx";
        return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
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

    private static void FormatearHeader(IXLRange range, XLColor bg, XLColor fg)
    {
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = bg;
        range.Style.Font.FontColor = fg;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
}
