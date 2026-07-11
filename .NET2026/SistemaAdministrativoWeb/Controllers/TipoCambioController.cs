using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("TIPOCAMBIO")]
public class TipoCambioController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    ITipoCambioRepository tipoCambioRepository,
    IMonedaRepository monedaRepository,
    ITipoCambioSyncService tipoCambioSyncService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var tiposCambio = await tipoCambioRepository.ListarPorCuentaAdministradoraAsync(contexto.IdCuentaAdministradora, anioTrabajo, mesTrabajo, cancellationToken);
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();

        return View(ConstruirViewModel(
            currentCompanyAccessor.EmpresaId!.Value,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            contexto.IdCuentaAdministradora,
            anioTrabajo,
            mesTrabajo,
            monedas,
            tiposCambio,
            null));
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idTipoCambio, string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, idTipoCambio, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(TipoCambioIndexViewModel model, string? periodo = null, CancellationToken cancellationToken = default)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;
        var formulario = model.Formulario ?? new TipoCambioFormViewModel();
        NormalizarFormulario(formulario);
        RevalidarFormulario(formulario, nameof(TipoCambioIndexViewModel.Formulario));

        var periodoTrabajo = NormalizarPeriodo(periodo);

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(contexto.IdCuentaAdministradora, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            await tipoCambioRepository.GuardarAsync(new GuardarTipoCambioRequest
            {
                IdTipoCambio = formulario.IdTipoCambio,
                IdCuentaAdministradora = contexto.IdCuentaAdministradora,
                Fecha = formulario.Fecha,
                IdMoneda = formulario.IdMoneda,
                Compra = formulario.Compra,
                Venta = formulario.Venta,
                CompraSbs = formulario.CompraSbs,
                VentaSbs = formulario.VentaSbs,
                Fuente = formulario.Fuente,
                Estado = formulario.Estado,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["TipoCambioOk"] = formulario.IdTipoCambio.HasValue
                ? "Tipo de cambio actualizado correctamente."
                : "Tipo de cambio registrado correctamente.";

            return RedirectToAction(nameof(Index), new
            {
                anio = formulario.Fecha.Year,
                mes = formulario.Fecha.Month
            });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirViewModelErrorAsync(contexto.IdCuentaAdministradora, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SincronizarPeriodoDesdeApi(short anio, byte mes, CancellationToken cancellationToken = default)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            var resultado = await tipoCambioSyncService.SincronizarPeriodoAsync(
                contexto.IdCuentaAdministradora,
                anio,
                mes,
                User.Identity?.Name,
                cancellationToken);

            TempData["TipoCambioOk"] = resultado.TotalSincronizados > 0
                ? $"Se sincronizaron {resultado.TotalSincronizados} tipos de cambio del periodo {anio:0000}{mes:00}."
                : $"La API no devolvio tipos de cambio para el periodo {anio:0000}{mes:00}.";
        }
        catch (Exception ex)
        {
            TempData["TipoCambioError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio, mes });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SincronizarFechaDesdeApi(TipoCambioIndexViewModel model, string? periodo = null, int? idTipoCambioActual = null, CancellationToken cancellationToken = default)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var formulario = model.Formulario ?? new TipoCambioFormViewModel();
        NormalizarFormulario(formulario);
        RevalidarFormulario(formulario, nameof(TipoCambioIndexViewModel.Formulario));

        if (!ModelState.IsValid)
        {
            var periodoInvalido = string.IsNullOrWhiteSpace(periodo)
                ? $"{formulario.Fecha.Year:0000}{formulario.Fecha.Month:00}"
                : NormalizarPeriodo(periodo);

            var modelConError = await ConstruirViewModelErrorAsync(contexto.IdCuentaAdministradora, periodoInvalido, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        var codigoMoneda = formulario.IdMoneda;
        var periodoTrabajo = string.IsNullOrWhiteSpace(periodo)
            ? $"{formulario.Fecha.Year:0000}{formulario.Fecha.Month:00}"
            : NormalizarPeriodo(periodo);

        try
        {
            var tipoCambio = await tipoCambioSyncService.SincronizarFechaAsync(
                contexto.IdCuentaAdministradora,
                formulario.Fecha,
                codigoMoneda,
                User.Identity?.Name,
                cancellationToken);

            if (tipoCambio is null)
            {
                TempData["TipoCambioError"] = "La API no devolvio un tipo de cambio para la fecha y moneda seleccionadas.";
                return idTipoCambioActual.HasValue
                    ? RedirectToAction(nameof(Editar), new { idTipoCambio = idTipoCambioActual.Value, periodo = periodoTrabajo })
                    : RedirectToAction(nameof(Registrar), new { periodo = periodoTrabajo });
            }

            TempData["TipoCambioOk"] = $"Se sincronizo el tipo de cambio del {tipoCambio.Fecha:dd/MM/yyyy} desde la API.";
            return RedirectToAction(nameof(Editar), new { idTipoCambio = tipoCambio.IdTipoCambio, periodo = periodoTrabajo });
        }
        catch (Exception ex)
        {
            TempData["TipoCambioError"] = ex.Message;
            return idTipoCambioActual.HasValue
                ? RedirectToAction(nameof(Editar), new { idTipoCambio = idTipoCambioActual.Value, periodo = periodoTrabajo })
                : RedirectToAction(nameof(Registrar), new { periodo = periodoTrabajo });
        }
    }

    [HttpGet]
    public async Task<IActionResult> ObtenerValorPorFecha(DateOnly? fecha = null, string? idMoneda = null, CancellationToken cancellationToken = default)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return Json(new { ok = false, mensaje = "No existe una empresa activa en la sesion." });
        }

        var codigoMonedaSolicitada = (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
        var codigoMoneda = NormalizarCodigoMoneda(codigoMonedaSolicitada);
        if (!fecha.HasValue)
        {
            return Json(new { ok = false, mensaje = "Ingrese una fecha valida para consultar el tipo de cambio." });
        }

        if (string.IsNullOrWhiteSpace(codigoMonedaSolicitada))
        {
            return Json(new { ok = false, mensaje = "Seleccione una moneda valida para consultar el tipo de cambio." });
        }

        if (codigoMoneda.Length != 3)
        {
            return Json(new { ok = false, mensaje = "Seleccione una moneda valida para consultar el tipo de cambio." });
        }

        try
        {
            var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(
                contexto.IdCuentaAdministradora,
                fecha.Value,
                codigoMoneda,
                cancellationToken);

            if (tipoCambio is null)
            {
                tipoCambio = await tipoCambioSyncService.SincronizarFechaAsync(
                    contexto.IdCuentaAdministradora,
                    fecha.Value,
                    codigoMoneda,
                    User.Identity?.Name,
                    cancellationToken);
            }

            if (tipoCambio is null)
            {
                return Json(new
                {
                    ok = true,
                    encontrado = false,
                    mensaje = "La API no devolvio un tipo de cambio para la fecha y moneda seleccionadas."
                });
            }

            return Json(new
            {
                ok = true,
                encontrado = true,
                tipoCambio = tipoCambio.Venta,
                compra = tipoCambio.Compra,
                venta = tipoCambio.Venta,
                compraSbs = tipoCambio.CompraSbs,
                ventaSbs = tipoCambio.VentaSbs,
                fuente = tipoCambio.Fuente,
                fecha = tipoCambio.Fecha.ToString("yyyy-MM-dd"),
                idMoneda = tipoCambio.IdMoneda,
                monedaSolicitada = codigoMonedaSolicitada
            });
        }
        catch (Exception ex)
        {
            return Json(new
            {
                ok = false,
                mensaje = ex.Message
            });
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(string? periodo, int? idTipoCambio, CancellationToken cancellationToken)
    {
        var contexto = await ObtenerContextoAsync(cancellationToken);
        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var periodoTrabajo = NormalizarPeriodo(periodo);
        var anioTrabajo = short.Parse(periodoTrabajo[..4]);
        var mesTrabajo = byte.Parse(periodoTrabajo[4..]);
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposCambio = await tipoCambioRepository.ListarPorCuentaAdministradoraAsync(contexto.IdCuentaAdministradora, anioTrabajo, mesTrabajo, cancellationToken);
        var tipoCambioEditar = idTipoCambio.HasValue
            ? await tipoCambioRepository.ObtenerAsync(idTipoCambio.Value, contexto.IdCuentaAdministradora, cancellationToken)
            : null;

        return View("Formulario", ConstruirViewModel(
            currentCompanyAccessor.EmpresaId!.Value,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            contexto.IdCuentaAdministradora,
            anioTrabajo,
            mesTrabajo,
            monedas,
            tiposCambio,
            tipoCambioEditar));
    }

    private async Task<TipoCambioIndexViewModel> ConstruirViewModelErrorAsync(int idCuentaAdministradora, string periodo, TipoCambioFormViewModel formulario, CancellationToken cancellationToken)
    {
        var anioTrabajo = short.Parse(periodo[..4]);
        var mesTrabajo = byte.Parse(periodo[4..]);
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposCambio = await tipoCambioRepository.ListarPorCuentaAdministradoraAsync(idCuentaAdministradora, anioTrabajo, mesTrabajo, cancellationToken);

        var model = ConstruirViewModel(
            currentCompanyAccessor.EmpresaId ?? 0,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            idCuentaAdministradora,
            anioTrabajo,
            mesTrabajo,
            monedas,
            tiposCambio,
            null);

        model.Formulario = formulario;
        return model;
    }

    private async Task<ContextoSuscripcionEmpresaDto?> ObtenerContextoAsync(CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return null;
        }

        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
        if (contexto is null)
        {
            TempData["TipoCambioError"] = "No se pudo resolver la cuenta administradora de la empresa activa.";
        }

        return contexto;
    }

    private static void NormalizarFormulario(TipoCambioFormViewModel formulario)
    {
        formulario.IdMoneda = NormalizarCodigoMoneda(formulario.IdMoneda);
        formulario.Fuente = (formulario.Fuente ?? string.Empty).Trim().ToUpperInvariant();
    }

    private void RevalidarFormulario(TipoCambioFormViewModel formulario, string prefijo)
    {
        ModelState.ClearValidationState(prefijo);
        ActualizarValorCampo(ModelState, prefijo, nameof(TipoCambioFormViewModel.IdMoneda), formulario.IdMoneda);
        ActualizarValorCampo(ModelState, prefijo, nameof(TipoCambioFormViewModel.Fuente), formulario.Fuente);
        TryValidateModel(formulario, prefijo);
    }

    private static void ActualizarValorCampo(ModelStateDictionary modelState, string prefijo, string nombreCampo, string? valor)
    {
        foreach (var clave in modelState.Keys
                     .Where(x => string.Equals(x, $"{prefijo}.{nombreCampo}", StringComparison.OrdinalIgnoreCase)
                              || x.EndsWith($".{nombreCampo}", StringComparison.OrdinalIgnoreCase)
                              || string.Equals(x, nombreCampo, StringComparison.OrdinalIgnoreCase))
                     .ToList())
        {
            modelState.Remove(clave);
            modelState.SetModelValue(clave, valor ?? string.Empty, valor ?? string.Empty);
        }
    }

    private static string NormalizarCodigoMoneda(string? idMoneda)
    {
        var valor = (idMoneda ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(valor))
        {
            return string.Empty;
        }

        var separadores = new[] { " - ", "-", " " };
        foreach (var separador in separadores)
        {
            var indice = valor.IndexOf(separador, StringComparison.Ordinal);
            if (indice > 0)
            {
                valor = valor[..indice].Trim();
                break;
            }
        }

        return valor.Length > 3 ? valor[..3] : valor;
    }

    private static TipoCambioIndexViewModel ConstruirViewModel(
        int idEmpresa,
        string empresaNombre,
        int idCuentaAdministradora,
        short anio,
        byte mes,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<TipoCambioDto> tiposCambio,
        TipoCambioDto? tipoCambioEditar)
    {
        var items = tiposCambio
            .OrderByDescending(x => x.Fecha)
            .ThenBy(x => x.IdMoneda)
            .Select(x => new TipoCambioItemViewModel
            {
                IdTipoCambio = x.IdTipoCambio,
                Fecha = x.Fecha,
                IdMoneda = x.IdMoneda,
                Compra = x.Compra,
                Venta = x.Venta,
                CompraSbs = x.CompraSbs,
                VentaSbs = x.VentaSbs,
                Fuente = x.Fuente,
                Estado = x.Estado
            })
            .ToList();

        return new TipoCambioIndexViewModel
        {
            IdEmpresa = idEmpresa,
            EmpresaNombre = empresaNombre,
            IdCuentaAdministradora = idCuentaAdministradora,
            PeriodoConsulta = $"{anio:0000}{mes:00}",
            AnioSeleccionado = anio,
            MesSeleccionado = mes,
            TotalTiposCambio = items.Count,
            TotalActivos = items.Count(x => x.Estado),
            TotalInactivos = items.Count(x => !x.Estado),
            TiposCambio = items,
            Monedas = monedas.ToList(),
            Fuentes =
            [
                new OpcionCatalogoViewModel { Valor = "MANUAL", Texto = "MANUAL" },
                new OpcionCatalogoViewModel { Valor = "SUNAT", Texto = "SUNAT" },
                new OpcionCatalogoViewModel { Valor = "API", Texto = "API" }
            ],
            AniosDisponibles = Enumerable.Range(anio - 5, 11).ToList(),
            MesesDisponibles = Enumerable.Range(1, 12)
                .Select(x => new MesOpcionViewModel
                {
                    Valor = (byte)x,
                    Nombre = new DateTime(2000, x, 1).ToString("MMMM")
                })
                .ToList(),
            Formulario = tipoCambioEditar is null
                ? new TipoCambioFormViewModel
                {
                    Fecha = new DateOnly(anio, mes, 1)
                }
                : new TipoCambioFormViewModel
                {
                    IdTipoCambio = tipoCambioEditar.IdTipoCambio,
                    Fecha = tipoCambioEditar.Fecha,
                    IdMoneda = tipoCambioEditar.IdMoneda,
                    Compra = tipoCambioEditar.Compra,
                    Venta = tipoCambioEditar.Venta,
                    CompraSbs = tipoCambioEditar.CompraSbs,
                    VentaSbs = tipoCambioEditar.VentaSbs,
                    Fuente = tipoCambioEditar.Fuente,
                    Estado = tipoCambioEditar.Estado
                }
        };
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
}
