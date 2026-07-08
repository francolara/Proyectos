using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Data;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class AplicacionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPeriodoContableService periodoContableService,
    IAplicacionNotaCreditoRepository aplicacionRepository,
    IPersonaRepository personaRepository) : Controller
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

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var aplicaciones = await aplicacionRepository.ListarPaginadoPorEmpresaAsync(
            empresaId,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            pagina,
            TamanoPagina,
            cancellationToken);

        var model = new AplicacionNotaCreditoIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            PeriodoConsulta = $"{anioTrabajo:0000}{mesTrabajo:00}",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TotalAplicaciones = aplicaciones.TotalRecords,
            TotalImporteAplicado = aplicaciones.Items.Sum(x => x.ImporteAplicado),
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).ToList(),
            MesesDisponibles = ConstruirMeses(),
            Aplicaciones = aplicaciones.Items.Select(x => new AplicacionNotaCreditoResumenItemViewModel
            {
                IdAplicacionNotaCredito = x.IdAplicacionNotaCredito,
                ModuloOperacion = x.ModuloOperacion,
                TipoPersonaTexto = x.TipoPersonaTexto,
                NombrePersona = x.NombrePersona,
                NumeroDocumentoPersona = x.NumeroDocumentoPersona,
                FechaAplicacion = x.FechaAplicacion,
                CodigoMoneda = x.CodigoMoneda,
                ImporteAplicado = x.ImporteAplicado,
                IdAsiento = x.IdAsiento,
                NumeroAsiento = x.NumeroAsiento,
                Glosa = x.Glosa,
                TipoComprobanteAplicado = x.TipoComprobanteAplicado,
                DescripcionTipoComprobanteAplicado = x.DescripcionTipoComprobanteAplicado,
                SerieAplicado = x.SerieAplicado,
                NumeroAplicado = x.NumeroAplicado,
                TipoComprobanteNc = x.TipoComprobanteNc,
                DescripcionTipoComprobanteNc = x.DescripcionTipoComprobanteNc,
                SerieNc = x.SerieNc,
                NumeroNc = x.NumeroNc
            }).ToList(),
            Paginacion = new PaginacionViewModel
            {
                PaginaActual = pagina,
                TamanoPagina = TamanoPagina,
                TotalRegistros = aplicaciones.TotalRecords
            }
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(short? anio = null, byte? mes = null, string? tipoPersona = null, int? idPersona = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["AplicacionError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }

        var formulario = new AplicacionNotaCreditoFormViewModel
        {
            TipoPersona = NormalizarTipoPersona(tipoPersona),
            IdPersona = idPersona,
            FechaAplicacion = new DateOnly(anioTrabajo, mesTrabajo, 1)
        };

        var model = await ConstruirFormularioAsync(formulario, anioTrabajo, mesTrabajo, cancellationToken);
        return View("Formulario", model);
    }

    [HttpGet]
    public async Task<IActionResult> BuscarPersonasAyuda(string? tipoPersona = null, string? textoBusqueda = null, int numeroPagina = 1, int tamanoPagina = 20, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "No existe una empresa activa en la sesion." });
        }

        var tipoPersonaTrabajo = NormalizarTipoPersona(tipoPersona);
        var resultado = await personaRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            textoBusqueda,
            null,
            tipoPersonaTrabajo == "C",
            tipoPersonaTrabajo == "P",
            numeroPagina <= 0 ? 1 : numeroPagina,
            tamanoPagina <= 0 ? 20 : tamanoPagina,
            cancellationToken);

        return Json(new
        {
            ok = true,
            items = resultado.Items.Select(x => new
            {
                idPersona = x.IdPersona,
                numeroDocumento = x.NumeroDocumento,
                nombreCompleto = x.NombreCompleto,
                tipoPersona = x.TipoPersona,
                nombreTipoDocumento = x.NombreTipoDocumento,
                esCliente = x.EsCliente,
                esProveedor = x.EsProveedor
            }),
            totalRegistros = resultado.TotalRecords
        });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(AplicacionNotaCreditoFormViewModel formulario, short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);
        if (await periodoContableService.EstaCerradoAsync(
                currentCompanyAccessor.EmpresaId.Value,
                (short)formulario.FechaAplicacion.Year,
                (byte)formulario.FechaAplicacion.Month,
                cancellationToken))
        {
            ModelState.AddModelError(
                string.Empty,
                periodoContableService.ConstruirMensajeBloqueo(
                    (short)formulario.FechaAplicacion.Year,
                    (byte)formulario.FechaAplicacion.Month));
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var model = await ConstruirFormularioAsync(formulario, anioTrabajo, mesTrabajo, cancellationToken);

        if (!ModelState.IsValid)
        {
            return View("Formulario", model);
        }

        var comprobante = model.ComprobantesPendientes.FirstOrDefault(x => x.IdRegistro == formulario.IdRegistroComprobante);
        var notaCredito = model.NotasCreditoPendientes.FirstOrDefault(x => x.IdRegistro == formulario.IdRegistroNotaCredito);

        if (comprobante is null)
        {
            ModelState.AddModelError(nameof(formulario.IdRegistroComprobante), "Seleccione un comprobante pendiente valido.");
        }

        if (notaCredito is null)
        {
            ModelState.AddModelError(nameof(formulario.IdRegistroNotaCredito), "Seleccione una nota de credito valida.");
        }

        if (comprobante is not null && notaCredito is not null)
        {
            if (!string.Equals(comprobante.CodigoMoneda, notaCredito.CodigoMoneda, StringComparison.OrdinalIgnoreCase))
            {
                ModelState.AddModelError(string.Empty, "El comprobante y la nota de credito deben estar en la misma moneda.");
            }

            var importeMaximo = Math.Min(comprobante.Saldo, notaCredito.Saldo);
            if (formulario.ImporteAplicado > importeMaximo)
            {
                ModelState.AddModelError(nameof(formulario.ImporteAplicado), $"El importe aplicado no puede exceder {importeMaximo:0.00}.");
            }

            formulario.IdMoneda = comprobante.IdMoneda;
            formulario.MonedaTexto = comprobante.CodigoMoneda;
        }

        if (!ModelState.IsValid)
        {
            return View("Formulario", model);
        }

        var moduloOperacion = formulario.TipoPersona == "P" ? "COM" : "VEN";
        try
        {
            var resultado = await aplicacionRepository.GuardarAsync(new GuardarAplicacionNotaCreditoRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId!.Value,
                ModuloOperacion = moduloOperacion,
                IdPersona = formulario.IdPersona!.Value,
                FechaAplicacion = formulario.FechaAplicacion,
                TipoCambio = formulario.TipoCambio,
                IdRegistroComprobante = formulario.IdRegistroComprobante!.Value,
                IdRegistroNotaCredito = formulario.IdRegistroNotaCredito!.Value,
                ImporteAplicado = formulario.ImporteAplicado,
                Glosa = formulario.Glosa,
                Observacion = formulario.Observacion,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["AplicacionOk"] = $"Aplicacion registrada correctamente. Asiento generado: {resultado.NumeroAsiento?.ToString() ?? "-"} .";
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            return View("Formulario", model);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idAplicacionNotaCredito, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        if (await periodoContableService.EstaCerradoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken))
        {
            TempData["AplicacionError"] = periodoContableService.ConstruirMensajeBloqueo(anioTrabajo, mesTrabajo);
            return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
        }

        try
        {
            await aplicacionRepository.EliminarAsync(idAplicacionNotaCredito, currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            TempData["AplicacionOk"] = "Aplicacion eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["AplicacionError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
    }

    private async Task<AplicacionNotaCreditoIndexViewModel> ConstruirFormularioAsync(AplicacionNotaCreditoFormViewModel formulario, short anio, byte mes, CancellationToken cancellationToken)
    {
        var empresaId = currentCompanyAccessor.EmpresaId!.Value;
        var tipoPersona = NormalizarTipoPersona(formulario.TipoPersona);
        formulario.TipoPersona = tipoPersona;

        PersonaDetalleDto? persona = null;
        if (formulario.IdPersona.HasValue && formulario.IdPersona.Value > 0)
        {
            persona = await personaRepository.ObtenerPorIdAsync(empresaId, formulario.IdPersona.Value, cancellationToken);
        }

        var moduloOperacion = tipoPersona == "P" ? "COM" : "VEN";
        var pendientes = persona is not null
            ? await aplicacionRepository.ListarPendientesPorPersonaAsync(empresaId, moduloOperacion, persona.IdPersona, cancellationToken)
            : [];

        if (formulario.IdPersona.HasValue && persona is null)
        {
            ModelState.AddModelError(nameof(formulario.IdPersona), "La persona seleccionada no corresponde al tipo indicado.");
        }

        if (persona is not null)
        {
            var tipoValido = tipoPersona == "C" ? persona.EsCliente : persona.EsProveedor;
            if (!tipoValido)
            {
                ModelState.AddModelError(nameof(formulario.IdPersona), tipoPersona == "C"
                    ? "La persona seleccionada no esta registrada como cliente."
                    : "La persona seleccionada no esta registrada como proveedor.");
            }

            formulario.PersonaTexto = persona.NombreCompleto;
            formulario.NumeroDocumentoPersona = persona.NumeroDocumento;
        }

        var comprobanteSeleccionado = formulario.IdRegistroComprobante.HasValue
            ? pendientes.FirstOrDefault(x => !x.EsNotaCredito && x.IdRegistro == formulario.IdRegistroComprobante.Value)
            : null;
        var notaCreditoSeleccionada = formulario.IdRegistroNotaCredito.HasValue
            ? pendientes.FirstOrDefault(x => x.EsNotaCredito && x.IdRegistro == formulario.IdRegistroNotaCredito.Value)
            : null;

        var monedaSeleccionada = comprobanteSeleccionado ?? notaCreditoSeleccionada;
        formulario.IdMoneda = monedaSeleccionada?.IdMoneda;
        formulario.MonedaTexto = monedaSeleccionada?.CodigoMoneda ?? string.Empty;

        if (formulario.TipoCambio <= 0)
        {
            formulario.TipoCambio = comprobanteSeleccionado?.TipoCambio
                ?? notaCreditoSeleccionada?.TipoCambio
                ?? 1m;
        }

        return new AplicacionNotaCreditoIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            PeriodoConsulta = $"{anio:0000}{mes:00}",
            AnioSeleccionado = anio,
            MesSeleccionado = mes,
            AniosDisponibles = Enumerable.Range(anio - 5, 11).ToList(),
            MesesDisponibles = ConstruirMeses(),
            ComprobantesPendientes = pendientes
                .Where(x => !x.EsNotaCredito)
                .Select(MapearPendiente)
                .ToList(),
            NotasCreditoPendientes = pendientes
                .Where(x => x.EsNotaCredito)
                .Select(MapearPendiente)
                .ToList(),
            Formulario = formulario
        };
    }

    private static AplicacionNotaCreditoPendienteItemViewModel MapearPendiente(AplicacionNotaCreditoPendienteDto x)
    {
        return new AplicacionNotaCreditoPendienteItemViewModel
        {
            IdRegistro = x.IdRegistro,
            IdMoneda = x.IdMoneda,
            FechaEmision = x.FechaEmision,
            TipoComprobante = x.TipoComprobante,
            DescripcionTipoComprobante = x.DescripcionTipoComprobante,
            Serie = x.Serie,
            Numero = x.Numero,
            CodigoMoneda = x.CodigoMoneda,
            TipoCambio = x.TipoCambio,
            ImporteTotal = x.ImporteTotal,
            Saldo = x.Saldo,
            EscenarioOperacion = x.EscenarioOperacion,
            Observacion = x.Observacion
        };
    }

    private static string NormalizarTipoPersona(string? tipoPersona)
    {
        return string.Equals((tipoPersona ?? string.Empty).Trim(), "P", StringComparison.OrdinalIgnoreCase) ? "P" : "C";
    }

    private static void NormalizarFormulario(AplicacionNotaCreditoFormViewModel formulario)
    {
        formulario.TipoPersona = NormalizarTipoPersona(formulario.TipoPersona);
        formulario.PersonaTexto = (formulario.PersonaTexto ?? string.Empty).Trim();
        formulario.NumeroDocumentoPersona = (formulario.NumeroDocumentoPersona ?? string.Empty).Trim();
        formulario.MonedaTexto = (formulario.MonedaTexto ?? string.Empty).Trim();
        formulario.Glosa = (formulario.Glosa ?? string.Empty).Trim();
        formulario.Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim();
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var hoy = DateTime.Today;
        return (anio ?? (short)hoy.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)hoy.Month);
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
}
