using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Plataforma;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize(Roles = "SuperAdmin")]
public class PlataformaController(ICuentaAdministradoraRepository cuentaAdministradoraRepository) : Controller
{
    private const int TamanoPaginaSuscriptores = 10;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda, string? estadoFiltro, int pagina = 1, CancellationToken cancellationToken = default)
    {
        ViewData["PlatformShell"] = true;

        var estadoNormalizado = NormalizarFiltroEstado(estadoFiltro);
        var textoNormalizado = (textoBusqueda ?? string.Empty).Trim();
        var resultado = await cuentaAdministradoraRepository.ListarCuentasSuscripcionPaginadasAsync(
            textoNormalizado,
            estadoNormalizado,
            Math.Max(1, pagina),
            TamanoPaginaSuscriptores,
            cancellationToken);
        var cuentas = await ConstruirCuentasSuscripcionAsync(resultado.Cuentas, cancellationToken);

        var model = new PlataformaIndexViewModel
        {
            TotalCuentas = resultado.TotalCuentas,
            CuentasActivas = resultado.CuentasActivas,
            CuentasEnPrueba = resultado.CuentasEnPrueba,
            CuentasSuspendidasOBaja = resultado.CuentasSuspendidasOBaja,
            CobrosRegistrados = resultado.CobrosRegistrados,
            CobrosPendientesAplicacion = resultado.CobrosPendientesAplicacion,
            MontoCobradoMes = resultado.MontoCobradoMes,
            TextoBusqueda = textoNormalizado,
            EstadoFiltro = estadoNormalizado,
            PaginaActual = resultado.PaginaActual,
            TamanoPagina = resultado.TamanoPagina,
            TotalRegistrosFiltrados = resultado.TotalFiltrado,
            TotalPaginas = resultado.TotalPaginas,
            Cuentas = cuentas
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Cobros(string? textoBusqueda, string? estadoPagoFiltro, CancellationToken cancellationToken)
    {
        ViewData["PlatformShell"] = true;

        var cuentas = await ConstruirCuentasSuscripcionAsync(cancellationToken);
        var textoNormalizado = (textoBusqueda ?? string.Empty).Trim();
        var estadoNormalizado = NormalizarFiltroEstadoCobro(estadoPagoFiltro);

        var cobros = cuentas
            .SelectMany(cuenta => cuenta.HistorialCobros.Select(pago => new PlataformaCobroItemViewModel
            {
                IdCuentaAdministradora = cuenta.IdCuentaAdministradora,
                IdCuentaAdministradoraSuscripcionPago = pago.IdCuentaAdministradoraSuscripcionPago,
                NombreCuenta = cuenta.NombreCuenta,
                CodigoCuenta = cuenta.CodigoCuenta,
                Contacto = cuenta.NombreCompleto ?? cuenta.Email ?? cuenta.CorreoPrincipal,
                TipoPago = pago.TipoPago,
                EstadoPago = pago.EstadoPago,
                Monto = pago.Monto,
                Moneda = pago.Moneda,
                FechaPago = pago.FechaPago,
                OperacionNumero = pago.OperacionNumero,
                EntidadFinanciera = pago.EntidadFinanciera,
                ReferenciaExterna = pago.ReferenciaExterna,
                ProveedorPasarela = pago.ProveedorPasarela,
                EstadoPasarela = pago.EstadoPasarela,
                AccionAplicacion = pago.AccionAplicacion,
                AplicarAlConfirmar = pago.AplicarAlConfirmar,
                AplicadoSuscripcion = pago.AplicadoSuscripcion,
                TipoCobroObjetivo = pago.TipoCobroObjetivo,
                FechaInicioPlanObjetivo = pago.FechaInicioPlanObjetivo,
                Observacion = pago.Observacion,
                FechaRegistro = pago.FechaRegistro
            }))
            .OrderByDescending(x => x.FechaPago)
            .ToList();

        if (!string.IsNullOrWhiteSpace(textoNormalizado))
        {
            cobros = cobros
                .Where(x =>
                    ContieneTexto(x.NombreCuenta, textoNormalizado) ||
                    ContieneTexto(x.CodigoCuenta, textoNormalizado) ||
                    ContieneTexto(x.Contacto, textoNormalizado) ||
                    ContieneTexto(x.OperacionNumero, textoNormalizado) ||
                    ContieneTexto(x.ReferenciaExterna, textoNormalizado) ||
                    ContieneTexto(x.ProveedorPasarela, textoNormalizado))
                .ToList();
        }

        if (!string.Equals(estadoNormalizado, "TODOS", StringComparison.OrdinalIgnoreCase))
        {
            cobros = cobros
                .Where(x => string.Equals(x.EstadoPago, estadoNormalizado, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var model = new PlataformaCobrosViewModel
        {
            TextoBusqueda = textoNormalizado,
            EstadoPagoFiltro = estadoNormalizado,
            TotalCobros = cobros.Count,
            CobrosPagados = cobros.Count(x => string.Equals(x.EstadoPago, "PAGADO", StringComparison.OrdinalIgnoreCase)),
            CobrosPendientes = cobros.Count(x => string.Equals(x.EstadoPago, "PENDIENTE", StringComparison.OrdinalIgnoreCase)),
            TotalMontoPagado = cobros.Where(x => string.Equals(x.EstadoPago, "PAGADO", StringComparison.OrdinalIgnoreCase)).Sum(x => x.Monto),
            TotalMontoPendiente = cobros.Where(x => string.Equals(x.EstadoPago, "PENDIENTE", StringComparison.OrdinalIgnoreCase)).Sum(x => x.Monto),
            Cobros = cobros
        };

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ActualizarSuscripcion(ActualizarSuscripcionCuentaViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            TempData["SuperAdminError"] = "No se pudo actualizar la suscripcion.";
            return RedirectToAction(nameof(Index));
        }

        await cuentaAdministradoraRepository.ActualizarSuscripcionCuentaAsync(new ActualizarSuscripcionCuentaRequest
        {
            IdCuentaAdministradora = model.IdCuentaAdministradora,
            TipoPlan = model.TipoPlan.Trim().ToUpperInvariant(),
            EstadoSuscripcion = model.EstadoSuscripcion.Trim().ToUpperInvariant(),
            EsPrueba = model.EsPrueba,
            FechaInicioPrueba = model.FechaInicioPrueba,
            FechaFinPrueba = model.FechaFinPrueba,
            FechaInicioPlan = model.FechaInicioPlan,
            FechaFinPlan = model.FechaFinPlan,
            TipoCobro = string.IsNullOrWhiteSpace(model.TipoCobro) ? null : model.TipoCobro.Trim().ToUpperInvariant(),
            DiasGracia = model.DiasGracia <= 0 ? 5 : model.DiasGracia,
            EmpresasPermitidas = model.EmpresasPermitidas,
            UsuariosPermitidos = model.UsuariosPermitidos,
            Activo = model.Activo,
            EstadoCuenta = model.EstadoCuenta,
            Observacion = string.IsNullOrWhiteSpace(model.Observacion) ? null : model.Observacion.Trim(),
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuperAdminOk"] = "Suscripcion actualizada.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ActivarContrato(ActivarContratoCuentaViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            TempData["SuperAdminError"] = "No se pudo iniciar el contrato.";
            return RedirectToAction(nameof(Index));
        }

        var fechaFin = CalcularFechaFinContrato(model.TipoCobro, model.FechaInicioPlan);
        await cuentaAdministradoraRepository.ActivarContratoCuentaAsync(new ActivarContratoCuentaRequest
        {
            IdCuentaAdministradora = model.IdCuentaAdministradora,
            TipoPlan = model.TipoPlan.Trim().ToUpperInvariant(),
            TipoCobro = model.TipoCobro.Trim().ToUpperInvariant(),
            FechaInicioPlan = model.FechaInicioPlan,
            FechaFinPlan = fechaFin,
            DiasGracia = model.DiasGracia <= 0 ? 5 : model.DiasGracia,
            Observacion = string.IsNullOrWhiteSpace(model.Observacion) ? null : model.Observacion.Trim(),
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuperAdminOk"] = "Contrato iniciado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RegistrarCobro(RegistrarPagoSuscripcionCuentaViewModel model, CancellationToken cancellationToken)
    {
        if (!ModelState.IsValid)
        {
            TempData["SuperAdminError"] = "No se pudo registrar el cobro.";
            return RedirectToAction(nameof(Index));
        }

        var accionAplicacion = NormalizarAccionAplicacionCobro(model.AccionAplicacion);
        var aplicarAlConfirmar = model.AplicarAlConfirmar || !string.IsNullOrWhiteSpace(accionAplicacion);

        await cuentaAdministradoraRepository.RegistrarPagoSuscripcionCuentaAsync(new RegistrarPagoSuscripcionCuentaRequest
        {
            IdCuentaAdministradora = model.IdCuentaAdministradora,
            TipoPago = model.TipoPago.Trim().ToUpperInvariant(),
            EstadoPago = model.CobroConfirmado ? "PAGADO" : "PENDIENTE",
            Monto = model.Monto,
            Moneda = "PEN",
            FechaPago = model.FechaPago ?? DateTime.Now,
            FechaVencimiento = model.FechaVencimiento,
            OperacionNumero = string.IsNullOrWhiteSpace(model.OperacionNumero) ? null : model.OperacionNumero.Trim(),
            EntidadFinanciera = string.IsNullOrWhiteSpace(model.EntidadFinanciera) ? null : model.EntidadFinanciera.Trim(),
            ReferenciaExterna = string.IsNullOrWhiteSpace(model.ReferenciaExterna) ? null : model.ReferenciaExterna.Trim(),
            ProveedorPasarela = string.IsNullOrWhiteSpace(model.ProveedorPasarela) ? null : model.ProveedorPasarela.Trim(),
            TransaccionPasarelaId = string.IsNullOrWhiteSpace(model.TransaccionPasarelaId) ? null : model.TransaccionPasarelaId.Trim(),
            PagoPasarelaId = string.IsNullOrWhiteSpace(model.PagoPasarelaId) ? null : model.PagoPasarelaId.Trim(),
            EstadoPasarela = string.IsNullOrWhiteSpace(model.EstadoPasarela) ? null : model.EstadoPasarela.Trim().ToUpperInvariant(),
            PayloadPasarela = string.IsNullOrWhiteSpace(model.PayloadPasarela) ? null : model.PayloadPasarela.Trim(),
            Observacion = string.IsNullOrWhiteSpace(model.Observacion) ? null : model.Observacion.Trim(),
            AccionAplicacion = accionAplicacion,
            AplicarAlConfirmar = aplicarAlConfirmar,
            TipoCobroObjetivo = string.IsNullOrWhiteSpace(model.TipoCobroObjetivo) ? null : model.TipoCobroObjetivo.Trim().ToUpperInvariant(),
            FechaInicioPlanObjetivo = model.FechaInicioPlanObjetivo,
            DiasGraciaObjetivo = model.DiasGraciaObjetivo,
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuperAdminOk"] = "Cobro registrado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ConfirmarCobro(int idCuentaAdministradora, int idCuentaAdministradoraSuscripcionPago, CancellationToken cancellationToken)
    {
        await cuentaAdministradoraRepository.ConfirmarPagoSuscripcionCuentaAsync(
            idCuentaAdministradora,
            idCuentaAdministradoraSuscripcionPago,
            User.Identity?.Name,
            cancellationToken);

        TempData["SuperAdminOk"] = "Cobro confirmado y aplicado.";
        return RedirectToAction(nameof(Index));
    }

    private static DateOnly CalcularFechaFinContrato(string? tipoCobro, DateOnly fechaInicio)
    {
        return (tipoCobro ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "TRIMESTRAL" => fechaInicio.AddMonths(3).AddDays(-1),
            "SEMESTRAL" => fechaInicio.AddMonths(6).AddDays(-1),
            "ANUAL" => fechaInicio.AddYears(1).AddDays(-1),
            _ => fechaInicio.AddMonths(1).AddDays(-1)
        };
    }

    private static string? NormalizarAccionAplicacionCobro(string? accionAplicacion)
    {
        var valor = (accionAplicacion ?? string.Empty).Trim().ToUpperInvariant();
        return valor switch
        {
            "ACTIVAR_CONTRATO" => valor,
            _ => null
        };
    }

    private async Task<List<CuentaSuscripcionViewModel>> ConstruirCuentasSuscripcionAsync(CancellationToken cancellationToken)
    {
        var cuentas = await cuentaAdministradoraRepository.ListarCuentasSuscripcionAsync(cancellationToken);
        return await ConstruirCuentasSuscripcionAsync(cuentas, cancellationToken);
    }

    private async Task<List<CuentaSuscripcionViewModel>> ConstruirCuentasSuscripcionAsync(
        IEnumerable<CuentaSuscripcionResumenDto> cuentas,
        CancellationToken cancellationToken)
    {
        var cuentasViewModel = new List<CuentaSuscripcionViewModel>();

        foreach (var x in cuentas)
        {
            var movimientos = await cuentaAdministradoraRepository.ListarMovimientosSuscripcionCuentaAsync(x.IdCuentaAdministradora, 12, cancellationToken);
            var pagos = await cuentaAdministradoraRepository.ListarPagosSuscripcionCuentaAsync(x.IdCuentaAdministradora, 24, cancellationToken);

            cuentasViewModel.Add(new CuentaSuscripcionViewModel
            {
                IdCuentaAdministradora = x.IdCuentaAdministradora,
                CodigoCuenta = x.CodigoCuenta,
                NombreCuenta = x.NombreCuenta,
                CorreoPrincipal = x.CorreoPrincipal,
                TelefonoPrincipal = x.TelefonoPrincipal,
                NombreCompleto = x.NombreCompleto,
                Telefono = x.Telefono,
                Email = x.Email,
                CantidadEmpresas = x.CantidadEmpresas,
                IdEmpresaPrincipal = x.IdEmpresaPrincipal,
                CodigoEmpresaPrincipal = x.CodigoEmpresaPrincipal,
                RazonSocialEmpresaPrincipal = x.RazonSocialEmpresaPrincipal,
                NombreComercialEmpresaPrincipal = x.NombreComercialEmpresaPrincipal,
                RucEmpresaPrincipal = x.RucEmpresaPrincipal,
                TipoPlan = x.TipoPlan ?? "TRIAL",
                EstadoSuscripcion = x.EstadoSuscripcion ?? "TRIAL",
                EstadoAccesoEfectivo = ResolverEstadoAccesoEfectivo(x, DateOnly.FromDateTime(DateTime.Today)),
                EsPrueba = x.EsPrueba,
                FechaInicioPrueba = x.FechaInicioPrueba,
                FechaFinPrueba = x.FechaFinPrueba,
                FechaInicioPlan = x.FechaInicioPlan,
                FechaFinPlan = x.FechaFinPlan,
                TipoCobro = x.TipoCobro,
                DiasGracia = x.DiasGracia,
                FechaFinGracia = x.FechaFinGracia,
                EmpresasPermitidas = x.EmpresasPermitidas,
                UsuariosPermitidos = x.UsuariosPermitidos,
                Activo = x.Activo,
                EstadoCuenta = x.EstadoCuenta,
                Observacion = x.Observacion,
                HistorialComercial = movimientos.Select(MapearMovimiento).ToList(),
                HistorialCobros = pagos.Select(MapearPago).ToList()
            });
        }

        return cuentasViewModel;
    }

    private static string ResolverEstadoAccesoEfectivo(
        CuentaSuscripcionResumenDto cuenta,
        DateOnly fechaActual)
    {
        var estado = (cuenta.EstadoSuscripcion ?? string.Empty).Trim().ToUpperInvariant();
        var tipoPlan = (cuenta.TipoPlan ?? string.Empty).Trim().ToUpperInvariant();

        if (estado == "BAJA")
        {
            return "BAJA";
        }

        if (!cuenta.Activo || estado == "SUSPENDIDO")
        {
            return "SUSPENDIDO";
        }

        var esPrueba = cuenta.EsPrueba
            || tipoPlan is "TRIAL" or "GRATIS"
            || estado == "TRIAL";
        if (esPrueba)
        {
            return !cuenta.FechaFinPrueba.HasValue || fechaActual > cuenta.FechaFinPrueba.Value
                ? "VENCIDO"
                : "TRIAL";
        }

        if (!cuenta.FechaFinPlan.HasValue)
        {
            return "SIN_VIGENCIA";
        }

        if (fechaActual <= cuenta.FechaFinPlan.Value)
        {
            return "ACTIVO";
        }

        var fechaFinGracia = cuenta.FechaFinGracia
            ?? (cuenta.DiasGracia > 0
                ? cuenta.FechaFinPlan.Value.AddDays(cuenta.DiasGracia)
                : (DateOnly?)null);

        return fechaFinGracia.HasValue && fechaActual <= fechaFinGracia.Value
            ? "EN_GRACIA"
            : "VENCIDO";
    }

    private static CuentaSuscripcionMovimientoViewModel MapearMovimiento(CuentaSuscripcionMovimientoDto m)
    {
        return new CuentaSuscripcionMovimientoViewModel
        {
            IdCuentaAdministradoraSuscripcionMovimiento = m.IdCuentaAdministradoraSuscripcionMovimiento,
            TipoMovimiento = m.TipoMovimiento,
            TipoPlanAnterior = m.TipoPlanAnterior,
            TipoPlanNuevo = m.TipoPlanNuevo,
            EstadoSuscripcionAnterior = m.EstadoSuscripcionAnterior,
            EstadoSuscripcionNuevo = m.EstadoSuscripcionNuevo,
            EsPruebaAnterior = m.EsPruebaAnterior,
            EsPruebaNuevo = m.EsPruebaNuevo,
            TipoCobroAnterior = m.TipoCobroAnterior,
            TipoCobroNuevo = m.TipoCobroNuevo,
            FechaInicioReferencia = m.FechaInicioReferencia,
            FechaFinReferencia = m.FechaFinReferencia,
            DiasGracia = m.DiasGracia,
            DiasExtra = m.DiasExtra,
            Observacion = m.Observacion,
            FechaRegistro = m.FechaRegistro,
            UsuarioRegistro = m.UsuarioRegistro
        };
    }

    private static CuentaSuscripcionPagoViewModel MapearPago(CuentaSuscripcionPagoDto p)
    {
        return new CuentaSuscripcionPagoViewModel
        {
            IdCuentaAdministradoraSuscripcionPago = p.IdCuentaAdministradoraSuscripcionPago,
            TipoPago = p.TipoPago,
            EstadoPago = p.EstadoPago,
            Monto = p.Monto,
            Moneda = p.Moneda,
            FechaPago = p.FechaPago,
            FechaVencimiento = p.FechaVencimiento,
            OperacionNumero = p.OperacionNumero,
            EntidadFinanciera = p.EntidadFinanciera,
            ReferenciaExterna = p.ReferenciaExterna,
            ProveedorPasarela = p.ProveedorPasarela,
            EstadoPasarela = p.EstadoPasarela,
            AccionAplicacion = p.AccionAplicacion,
            AplicarAlConfirmar = p.AplicarAlConfirmar,
            AplicadoSuscripcion = p.AplicadoSuscripcion,
            FechaAplicacion = p.FechaAplicacion,
            UsuarioAplicacion = p.UsuarioAplicacion,
            TipoCobroObjetivo = p.TipoCobroObjetivo,
            FechaInicioPlanObjetivo = p.FechaInicioPlanObjetivo,
            DiasGraciaObjetivo = p.DiasGraciaObjetivo,
            Observacion = p.Observacion,
            FechaRegistro = p.FechaRegistro,
            UsuarioRegistro = p.UsuarioRegistro
        };
    }

    private static string NormalizarFiltroEstado(string? estadoFiltro)
    {
        var valor = (estadoFiltro ?? string.Empty).Trim().ToUpperInvariant();
        return valor switch
        {
            "ACTIVAS" => valor,
            "TRIAL" => valor,
            "SUSPENDIDAS" => valor,
            _ => "TODOS"
        };
    }

    private static string NormalizarFiltroEstadoCobro(string? estadoFiltro)
    {
        var valor = (estadoFiltro ?? string.Empty).Trim().ToUpperInvariant();
        return valor switch
        {
            "PAGADO" => valor,
            "PENDIENTE" => valor,
            _ => "TODOS"
        };
    }

    private static bool ContieneTexto(string? origen, string textoBusqueda)
    {
        return !string.IsNullOrWhiteSpace(origen)
               && origen.Contains(textoBusqueda, StringComparison.OrdinalIgnoreCase);
    }
}
