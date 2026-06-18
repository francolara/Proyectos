using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class AsientoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IAsientoRepository asientoRepository,
    IOrigenRepository origenRepository,
    IPlanCuentaRepository planCuentaRepository,
    IMonedaRepository monedaRepository) : Controller
{
    private const int TamanoPagina = 20;

    [HttpGet]
    public async Task<IActionResult> Index(short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var asientos = await asientoRepository.ListarPaginadoPorEmpresaAsync(empresaId, anioTrabajo, mesTrabajo, textoBusqueda, pagina, TamanoPagina, true, cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            origenes,
            monedas,
            cuentas,
            asientos.Items,
            null);
        model.TotalAsientos = asientos.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = asientos.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idAsiento, string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, idAsiento, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(AsientoFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);
        ValidarFormulario(formulario);

        var periodoTrabajo = $"{formulario.FechaAsiento.Year:0000}{formulario.FechaAsiento.Month:00}";

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var result = await asientoRepository.GuardarManualAsync(new GuardarAsientoManualRequest
            {
                IdAsiento = formulario.IdAsiento,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdOrigen = formulario.IdOrigen!.Value,
                FechaAsiento = formulario.FechaAsiento,
                Glosa = formulario.Glosa.Trim(),
                IdMoneda = formulario.IdMoneda!.Value,
                TipoCambio = formulario.TipoCambio,
                ReferenciaExterna = string.IsNullOrWhiteSpace(formulario.ReferenciaExterna) ? null : formulario.ReferenciaExterna.Trim(),
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarAsientoDetalleRequest
                    {
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        GlosaDetalle = string.IsNullOrWhiteSpace(x.GlosaDetalle) ? null : x.GlosaDetalle.Trim(),
                        Debe = x.Debe,
                        Haber = x.Haber,
                        ReferenciaLinea = string.IsNullOrWhiteSpace(x.ReferenciaLinea) ? null : x.ReferenciaLinea.Trim()
                    })
                    .ToList()
            }, cancellationToken);

            TempData["AsientoOk"] = $"Asiento {result.Periodo}-{result.NumeroAsiento} guardado correctamente.";
            return RedirectToAction(nameof(Index), new { anio = formulario.FechaAsiento.Year, mes = formulario.FechaAsiento.Month });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(string? periodo, int? idAsiento, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var asientos = await asientoRepository.ListarPorEmpresaAsync(empresaId, periodoTrabajo, true, cancellationToken);
        var asientoEditar = idAsiento.HasValue
            ? await asientoRepository.ObtenerAsync(idAsiento.Value, cancellationToken)
            : null;

        if (asientoEditar is not null && asientoEditar.IdEmpresa != empresaId)
        {
            asientoEditar = null;
        }

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            short.Parse(periodoTrabajo[..4]),
            byte.Parse(periodoTrabajo[4..]),
            null,
            origenes,
            monedas,
            cuentas,
            asientos,
            asientoEditar);

        return View("Formulario", model);
    }

    private async Task<AsientoIndexViewModel> ConstruirViewModelErrorAsync(int empresaId, string periodo, AsientoFormViewModel formulario, CancellationToken cancellationToken)
    {
        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.PermiteRegistroManual)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPorEmpresaAsync(empresaId, true, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var asientos = await asientoRepository.ListarPorEmpresaAsync(empresaId, periodo, true, cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodo,
            short.Parse(periodo[..4]),
            byte.Parse(periodo[4..]),
            null,
            origenes,
            monedas,
            cuentas,
            asientos,
            null);

        model.Formulario = formulario;
        return model;
    }

    private static void NormalizarFormulario(AsientoFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue
                     || !string.IsNullOrWhiteSpace(x.GlosaDetalle)
                     || !string.IsNullOrWhiteSpace(x.ReferenciaLinea)
                     || x.Debe > 0
                     || x.Haber > 0)
            .Select((x, index) =>
            {
                x.Item = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(AsientoFormViewModel formulario)
    {
        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos dos lineas en el asiento.");
            return;
        }

        if (formulario.Detalles.Count < 2)
        {
            ModelState.AddModelError(string.Empty, "El asiento debe tener al menos dos lineas.");
        }

        decimal totalDebe = 0;
        decimal totalHaber = 0;

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (!detalle.IdPlanCuenta.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione una cuenta.");
            }

            var tieneDebe = detalle.Debe > 0;
            var tieneHaber = detalle.Haber > 0;

            if (tieneDebe == tieneHaber)
            {
                ModelState.AddModelError($"{prefijo}.Debe", "La linea debe tener monto solo en Debe o solo en Haber.");
            }

            totalDebe += detalle.Debe;
            totalHaber += detalle.Haber;
        }

        if (totalDebe <= 0 || totalHaber <= 0)
        {
            ModelState.AddModelError(string.Empty, "El asiento debe tener importes positivos tanto en Debe como en Haber.");
        }

        if (decimal.Round(totalDebe, 2) != decimal.Round(totalHaber, 2))
        {
            ModelState.AddModelError(string.Empty, "El asiento no esta cuadrado. Debe y Haber deben ser iguales.");
        }
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var today = DateTime.Today;
        var anioTrabajo = anio ?? (short)today.Year;
        var mesTrabajo = mes is >= 1 and <= 12 ? mes.Value : (byte)today.Month;
        return (anioTrabajo, mesTrabajo);
    }

    private static string NormalizarPeriodo(string? periodo)
    {
        if (!string.IsNullOrWhiteSpace(periodo)
            && periodo.Length == 6
            && short.TryParse(periodo[..4], out var anio)
            && byte.TryParse(periodo[4..], out var mes)
            && mes is >= 1 and <= 12)
        {
            return $"{anio:0000}{mes:00}";
        }

        var (anioActual, mesActual) = NormalizarPeriodo(null, null);
        return $"{anioActual:0000}{mesActual:00}";
    }

    private static AsientoIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        string periodo,
        short anioSeleccionado,
        byte mesSeleccionado,
        string? textoBusqueda,
        IReadOnlyCollection<OrigenDto> origenes,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        IReadOnlyCollection<AsientoResumenDto> asientos,
        AsientoDto? asientoEditar)
    {
        var items = asientos
            .Select(x => new AsientoResumenItemViewModel
            {
                IdAsiento = x.IdAsiento,
                CodigoOrigen = x.CodigoOrigen,
                NombreOrigen = x.NombreOrigen,
                Periodo = x.Periodo,
                NumeroAsiento = x.NumeroAsiento,
                FechaAsiento = x.FechaAsiento,
                Glosa = x.Glosa,
                CodigoMoneda = x.CodigoMoneda,
                TipoCambio = x.TipoCambio,
                TotalDebe = x.TotalDebe,
                TotalHaber = x.TotalHaber,
                Estado = x.Estado
            })
            .ToList();

        return new AsientoIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TotalAsientos = items.Count,
            TotalDebePeriodo = items.Sum(x => x.TotalDebe),
            TotalHaberPeriodo = items.Sum(x => x.TotalHaber),
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            OrigenesManual = origenes.ToList(),
            Monedas = monedas.ToList(),
            CuentasMovimiento = cuentas.ToList(),
            Asientos = items,
            Formulario = asientoEditar is null
                ? new AsientoFormViewModel
                {
                    FechaAsiento = ParsePeriodo(periodo),
                    IdOrigen = origenes.FirstOrDefault()?.IdOrigen,
                    IdMoneda = monedas.OrderByDescending(x => x.EsMonedaBase).FirstOrDefault()?.IdMoneda
                }
                : new AsientoFormViewModel
                {
                    IdAsiento = asientoEditar.IdAsiento,
                    IdOrigen = asientoEditar.IdOrigen,
                    FechaAsiento = asientoEditar.FechaAsiento,
                    Glosa = asientoEditar.Glosa,
                    IdMoneda = asientoEditar.IdMoneda,
                    TipoCambio = asientoEditar.TipoCambio,
                    ReferenciaExterna = asientoEditar.ReferenciaExterna,
                    Observacion = asientoEditar.Observacion,
                    Detalles = asientoEditar.Detalles
                        .OrderBy(x => x.Item)
                        .Select(x => new AsientoDetalleFormViewModel
                        {
                            IdAsientoDetalle = x.IdAsientoDetalle,
                            Item = x.Item,
                            IdPlanCuenta = x.IdPlanCuenta,
                            GlosaDetalle = x.GlosaDetalle,
                            Debe = x.Debe,
                            Haber = x.Haber,
                            ReferenciaLinea = x.ReferenciaLinea
                        })
                        .ToList()
                }
        };
    }

    private static List<int> ConstruirAnios(short anioSeleccionado)
    {
        return Enumerable.Range(anioSeleccionado - 5, 11).ToList();
    }

    private static List<MesOpcionViewModel> ConstruirMeses()
    {
        return Enumerable.Range(1, 12)
            .Select(x => new MesOpcionViewModel
            {
                Valor = (byte)x,
                Nombre = new DateTime(2000, x, 1).ToString("MMMM")
            })
            .ToList();
    }

    private static DateOnly ParsePeriodo(string periodo)
    {
        if (periodo.Length == 6
            && int.TryParse(periodo[..4], out var year)
            && int.TryParse(periodo[4..], out var month)
            && month is >= 1 and <= 12)
        {
            return new DateOnly(year, month, 1);
        }

        return DateOnly.FromDateTime(DateTime.Today);
    }
}
