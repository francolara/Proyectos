using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class ProcesoController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPeriodoContableService periodoContableService,
    IConfiguracionContabilizacionRepository configuracionContabilizacionRepository,
    IDiferenciaCambioRepository diferenciaCambioRepository,
    IAjusteCuentaRepository ajusteCuentaRepository,
    IAperturaProcesoRepository aperturaProcesoRepository,
    ICierreProcesoRepository cierreProcesoRepository,
    ITipoCambioRepository tipoCambioRepository,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    IParametroEmpresaRepository parametroEmpresaRepository) : Controller
{
    [HttpGet]
    public async Task<IActionResult> CerrarPeriodo(short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var estado = await periodoContableService.ObtenerEstadoAsync(currentCompanyAccessor.EmpresaId.Value, anioTrabajo, mesTrabajo, cancellationToken);

        var model = new ProcesoCerrarPeriodoViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            PeriodoConsulta = $"{anioTrabajo:0000}{mesTrabajo:00}",
            AnioSeleccionado = anioTrabajo,
            MesSeleccionado = mesTrabajo,
            Cerrado = estado.Cerrado,
            FechaCierre = estado.FechaCierre,
            UsuarioCierre = estado.UsuarioCierre,
            FechaApertura = estado.FechaApertura,
            UsuarioApertura = estado.UsuarioApertura,
            AniosDisponibles = Enumerable.Range(anioTrabajo - 5, 11).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };

        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarCerrarPeriodo(short anio, byte mes, bool cerrado, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await periodoContableService.GuardarEstadoAsync(
                currentCompanyAccessor.EmpresaId.Value,
                anio,
                mes,
                cerrado,
                User.Identity?.Name,
                cancellationToken);

            TempData["ProcesoOk"] = cerrado
                ? $"El periodo {mes:00}/{anio:0000} se cerro correctamente."
                : $"El periodo {mes:00}/{anio:0000} se abrio correctamente.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(CerrarPeriodo), new { anio, mes });
    }

    [HttpGet]
    public async Task<IActionResult> DiferenciaCambio(short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var configuracion = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(idEmpresa, cancellationToken);
        var provisionDiferencia = configuracion.Provisiones
            .FirstOrDefault(x => string.Equals(x.ModuloOperacion, "DIF", StringComparison.OrdinalIgnoreCase) && x.Activo);
        var proceso = await diferenciaCambioRepository.ObtenerAsync(idEmpresa, periodo, cancellationToken);
        var (usaTipoCambioSbs, tipoCambioCompra, tipoCambioVenta) = await ResolverTipoCambioMensualAsync(idEmpresa, anioTrabajo, mesTrabajo, true, cancellationToken);

        return View(ConstruirViewModelDiferenciaCambio(
            anioTrabajo,
            mesTrabajo,
            provisionDiferencia,
            proceso,
            usaTipoCambioSbs,
            tipoCambioCompra,
            tipoCambioVenta));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GenerarDiferenciaCambio(short anio, byte mes, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodo = $"{anio:0000}{mes:00}";

        try
        {
            var proceso = await diferenciaCambioRepository.GenerarAsync(new GenerarDiferenciaCambioProcesoRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                Periodo = periodo,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["ProcesoOk"] = proceso.TotalAsientos > 0
                ? $"Se generaron {proceso.TotalAsientos} asientos de diferencia en cambio para el periodo {mes:00}/{anio:0000}."
                : $"El proceso de diferencia en cambio no genero asientos para el periodo {mes:00}/{anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(DiferenciaCambio), new { anio, mes });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarDiferenciaCambio(short anio, byte mes, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodo = $"{anio:0000}{mes:00}";

        try
        {
            await diferenciaCambioRepository.EliminarAsync(
                currentCompanyAccessor.EmpresaId.Value,
                periodo,
                User.Identity?.Name,
                cancellationToken);

            TempData["ProcesoOk"] = $"Se elimino la generacion de diferencia en cambio del periodo {mes:00}/{anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(DiferenciaCambio), new { anio, mes });
    }

    [HttpGet]
    public async Task<IActionResult> AjusteCuenta(short? anio = null, byte? mes = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);
        var periodo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var configuracion = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(idEmpresa, cancellationToken);
        var provisionAjuste = configuracion.Provisiones
            .FirstOrDefault(x => string.Equals(x.ModuloOperacion, "AJU", StringComparison.OrdinalIgnoreCase) && x.Activo);
        var proceso = await ajusteCuentaRepository.ObtenerAsync(idEmpresa, periodo, cancellationToken);
        var (usaTipoCambioSbs, tipoCambioCompra, tipoCambioVenta) = await ResolverTipoCambioMensualAsync(idEmpresa, anioTrabajo, mesTrabajo, false, cancellationToken);

        return View(ConstruirViewModelAjusteCuenta(
            anioTrabajo,
            mesTrabajo,
            provisionAjuste,
            proceso,
            usaTipoCambioSbs,
            tipoCambioCompra,
            tipoCambioVenta));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GenerarAjusteCuenta(short anio, byte mes, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodo = $"{anio:0000}{mes:00}";

        try
        {
            var proceso = await ajusteCuentaRepository.GenerarAsync(new GenerarAjusteCuentaProcesoRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                Periodo = periodo,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["ProcesoOk"] = proceso.TotalAsientos > 0
                ? $"Se generaron {proceso.TotalAsientos} asientos de ajuste de cuentas para el periodo {mes:00}/{anio:0000}."
                : $"El proceso de ajuste de cuentas no genero asientos para el periodo {mes:00}/{anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AjusteCuenta), new { anio, mes });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarAjusteCuenta(short anio, byte mes, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var periodo = $"{anio:0000}{mes:00}";

        try
        {
            await ajusteCuentaRepository.EliminarAsync(
                currentCompanyAccessor.EmpresaId.Value,
                periodo,
                User.Identity?.Name,
                cancellationToken);

            TempData["ProcesoOk"] = $"Se elimino la generacion de ajuste de cuentas del periodo {mes:00}/{anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AjusteCuenta), new { anio, mes });
    }

    [HttpGet]
    public async Task<IActionResult> AsientoApertura(short? anioApertura = null, byte? mesSaldoHasta = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var anioTrabajo = anioApertura ?? (short)DateTime.Today.Year;
        var mesSaldoTrabajo = mesSaldoHasta is <= 15 ? mesSaldoHasta.Value : (byte)12;
        var anioSaldo = (short)(anioTrabajo - 1);
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var configuracion = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(idEmpresa, cancellationToken);
        var provisionApertura = configuracion.Provisiones
            .FirstOrDefault(x => string.Equals(x.ModuloOperacion, "APR", StringComparison.OrdinalIgnoreCase) && x.Activo);
        var proceso = await aperturaProcesoRepository.ObtenerAsync(idEmpresa, anioTrabajo, cancellationToken);
        var (usaTipoCambioSbs, tipoCambioCompra, tipoCambioVenta) = await ResolverTipoCambioAperturaAsync(idEmpresa, anioSaldo, proceso, cancellationToken);

        return View(ConstruirViewModelApertura(
            anioTrabajo,
            anioSaldo,
            mesSaldoTrabajo,
            provisionApertura,
            proceso,
            usaTipoCambioSbs,
            tipoCambioCompra,
            tipoCambioVenta));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GenerarAsientoApertura(
        short anioApertura,
        byte mesSaldoHasta,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            var proceso = await aperturaProcesoRepository.GenerarAsync(new GenerarAperturaProcesoRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                AnioApertura = anioApertura,
                MesSaldoHasta = mesSaldoHasta,
                TipoCambioCompra = tipoCambioCompra,
                TipoCambioVenta = tipoCambioVenta,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["ProcesoOk"] = proceso.IdAsiento.HasValue
                ? $"Se genero el asiento de apertura Nro {proceso.NumeroAsiento} para el ejercicio {anioApertura:0000}."
                : $"El proceso de apertura no genero asiento para el ejercicio {anioApertura:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AsientoApertura), new { anioApertura, mesSaldoHasta });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarAsientoApertura(short anioApertura, byte mesSaldoHasta, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await aperturaProcesoRepository.EliminarAsync(
                currentCompanyAccessor.EmpresaId.Value,
                anioApertura,
                User.Identity?.Name,
                cancellationToken);

            TempData["ProcesoOk"] = $"Se elimino la generacion del asiento de apertura del ejercicio {anioApertura:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AsientoApertura), new { anioApertura, mesSaldoHasta });
    }

    [HttpGet]
    public async Task<IActionResult> AsientoCierre(short? anio = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var anioTrabajo = anio ?? (short)DateTime.Today.Year;
        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var configuracion = await configuracionContabilizacionRepository.ObtenerConfiguracionContableEmpresaAsync(idEmpresa, cancellationToken);
        var provisionCierre = configuracion.Provisiones
            .FirstOrDefault(x => string.Equals(x.ModuloOperacion, "CIE", StringComparison.OrdinalIgnoreCase) && x.Activo);
        var proceso = await cierreProcesoRepository.ObtenerAsync(idEmpresa, anioTrabajo, cancellationToken);
        var (usaTipoCambioSbs, tipoCambioCompra, tipoCambioVenta) = await ResolverTipoCambioCierreAsync(idEmpresa, anioTrabajo, proceso, cancellationToken);

        return View(ConstruirViewModelCierre(
            anioTrabajo,
            provisionCierre,
            proceso,
            usaTipoCambioSbs,
            tipoCambioCompra,
            tipoCambioVenta));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GenerarAsientoCierre(
        short anio,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta,
        bool procesarGananciasPerdidas = false,
        bool procesarInventarios = false,
        CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            var proceso = await cierreProcesoRepository.GenerarAsync(new GenerarCierreProcesoRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                Anio = anio,
                TipoCambioCompra = tipoCambioCompra,
                TipoCambioVenta = tipoCambioVenta,
                ProcesarGananciasPerdidas = procesarGananciasPerdidas,
                ProcesarInventarios = procesarInventarios,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["ProcesoOk"] = proceso.TotalAsientos > 0
                ? $"Se generaron {proceso.TotalAsientos} asientos de cierre para el ejercicio {anio:0000}."
                : $"El proceso de cierre no genero asientos para el ejercicio {anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AsientoCierre), new { anio });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> EliminarAsientoCierre(short anio, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await cierreProcesoRepository.EliminarAsync(
                currentCompanyAccessor.EmpresaId.Value,
                anio,
                User.Identity?.Name,
                cancellationToken);

            TempData["ProcesoOk"] = $"Se elimino la generacion del asiento de cierre del ejercicio {anio:0000}.";
        }
        catch (Exception ex)
        {
            TempData["ProcesoError"] = ex.Message;
        }

        return RedirectToAction(nameof(AsientoCierre), new { anio });
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var today = DateTime.Today;
        var anioTrabajo = anio ?? (short)today.Year;
        var mesTrabajo = mes is >= 1 and <= 12 ? mes.Value : (byte)today.Month;
        return (anioTrabajo, mesTrabajo);
    }

    private async Task<(bool UsaTipoCambioSbs, decimal TipoCambioCompra, decimal TipoCambioVenta)> ResolverTipoCambioMensualAsync(
        int idEmpresa,
        short anio,
        byte mes,
        bool usarSbsEnDiciembre,
        CancellationToken cancellationToken)
    {
        var usaTipoCambioSbs = false;
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, "CONTABLE", "TIPO_CAMBIO_SBS_CIERRE", 1, 10, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x => string.Equals(x.CodigoParametro, "TIPO_CAMBIO_SBS_CIERRE", StringComparison.OrdinalIgnoreCase) && x.Activo);
        if (parametro is not null)
        {
            usaTipoCambioSbs = usarSbsEnDiciembre
                && mes == 12
                && string.Equals(parametro.ValorParametro?.Trim(), "S", StringComparison.OrdinalIgnoreCase);
        }

        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(idEmpresa, cancellationToken);
        if (contexto is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        var fechaCierre = DateOnly.FromDateTime(DateTime.DaysInMonth(anio, mes) switch
        {
            var ultimoDia => new DateTime(anio, mes, ultimoDia)
        });
        var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(contexto.IdCuentaAdministradora, fechaCierre, "USD", cancellationToken);
        if (tipoCambio is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        return usaTipoCambioSbs
            ? (true, tipoCambio.CompraSbs, tipoCambio.VentaSbs)
            : (false, tipoCambio.Compra, tipoCambio.Venta);
    }

    private async Task<(bool UsaTipoCambioSbs, decimal TipoCambioCompra, decimal TipoCambioVenta)> ResolverTipoCambioAperturaAsync(
        int idEmpresa,
        short anioSaldo,
        AperturaProcesoDto? proceso,
        CancellationToken cancellationToken)
    {
        if (proceso is not null)
        {
            return (proceso.UsaTipoCambioSbs, proceso.TipoCambioCompra, proceso.TipoCambioVenta);
        }

        var usaTipoCambioSbs = false;
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, "CONTABLE", "TIPO_CAMBIO_SBS_CIERRE", 1, 10, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x => string.Equals(x.CodigoParametro, "TIPO_CAMBIO_SBS_CIERRE", StringComparison.OrdinalIgnoreCase) && x.Activo);
        if (parametro is not null)
        {
            usaTipoCambioSbs = string.Equals(parametro.ValorParametro?.Trim(), "S", StringComparison.OrdinalIgnoreCase);
        }

        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(idEmpresa, cancellationToken);
        if (contexto is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        var fechaCierre = new DateOnly(anioSaldo, 12, 31);
        var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(contexto.IdCuentaAdministradora, fechaCierre, "USD", cancellationToken);
        if (tipoCambio is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        return usaTipoCambioSbs
            ? (true, tipoCambio.CompraSbs, tipoCambio.VentaSbs)
            : (false, tipoCambio.Compra, tipoCambio.Venta);
    }

    private async Task<(bool UsaTipoCambioSbs, decimal TipoCambioCompra, decimal TipoCambioVenta)> ResolverTipoCambioCierreAsync(
        int idEmpresa,
        short anio,
        CierreProcesoDto? proceso,
        CancellationToken cancellationToken)
    {
        if (proceso is not null)
        {
            return (proceso.UsaTipoCambioSbs, proceso.TipoCambioCompra, proceso.TipoCambioVenta);
        }

        var usaTipoCambioSbs = false;
        var parametros = await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, "CONTABLE", "TIPO_CAMBIO_SBS_CIERRE", 1, 10, cancellationToken);
        var parametro = parametros.Items.FirstOrDefault(x => string.Equals(x.CodigoParametro, "TIPO_CAMBIO_SBS_CIERRE", StringComparison.OrdinalIgnoreCase) && x.Activo);
        if (parametro is not null)
        {
            usaTipoCambioSbs = string.Equals(parametro.ValorParametro?.Trim(), "S", StringComparison.OrdinalIgnoreCase);
        }

        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(idEmpresa, cancellationToken);
        if (contexto is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        var fechaCierre = new DateOnly(anio, 12, 31);
        var tipoCambio = await tipoCambioRepository.ObtenerPorFechaMonedaAsync(contexto.IdCuentaAdministradora, fechaCierre, "USD", cancellationToken);
        if (tipoCambio is null)
        {
            return (usaTipoCambioSbs, 0, 0);
        }

        return usaTipoCambioSbs
            ? (true, tipoCambio.CompraSbs, tipoCambio.VentaSbs)
            : (false, tipoCambio.Compra, tipoCambio.Venta);
    }

    private static List<MesOpcionViewModel> ListarMesesCalendario()
    {
        return Enumerable.Range(1, 12)
            .Select(x => new MesOpcionViewModel
            {
                Valor = (byte)x,
                Nombre = new DateTime(2000, x, 1).ToString("MMMM")
            })
            .ToList();
    }

    private static List<MesOpcionViewModel> ListarMesesContables()
    {
        string[] meses =
        [
            "Apertura",
            "Enero",
            "Febrero",
            "Marzo",
            "Abril",
            "Mayo",
            "Junio",
            "Julio",
            "Agosto",
            "Setiembre",
            "Octubre",
            "Noviembre",
            "Diciembre",
            "Ajustes y Liquidaciones",
            "Cierre de Ganancias y Perdidas",
            "Cierre de Inventarios"
        ];

        return meses
            .Select((nombre, index) => new MesOpcionViewModel
            {
                Valor = (byte)index,
                Nombre = nombre
            })
            .ToList();
    }

    private ProcesoDiferenciaCambioViewModel ConstruirViewModelDiferenciaCambio(
        short anio,
        byte mes,
        ConfiguracionContableProvisionDto? provisionDiferencia,
        DiferenciaCambioProcesoDto? proceso,
        bool usaTipoCambioSbs,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta)
    {
        return new ProcesoDiferenciaCambioViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId!.Value,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            PeriodoConsulta = $"{anio:0000}{mes:00}",
            AnioSeleccionado = anio,
            MesSeleccionado = mes,
            IdOrigenConfigurado = provisionDiferencia?.IdOrigen,
            OrigenConfiguradoTexto = provisionDiferencia is null
                ? "No configurado"
                : $"{provisionDiferencia.CodigoOrigen} - {provisionDiferencia.NombreOrigen}",
            PuedeGenerar = provisionDiferencia?.IdOrigen is not null,
            ProcesoGenerado = proceso is not null,
            FechaAsiento = proceso?.FechaAsiento,
            UsaTipoCambioSbs = proceso?.UsaTipoCambioSbs ?? usaTipoCambioSbs,
            TipoCambioCompra = proceso?.TipoCambioCompra ?? tipoCambioCompra,
            TipoCambioVenta = proceso?.TipoCambioVenta ?? tipoCambioVenta,
            TotalCuentas = proceso?.TotalCuentas ?? 0,
            TotalAsientos = proceso?.TotalAsientos ?? 0,
            TotalDebe = proceso?.TotalDebe ?? 0,
            TotalHaber = proceso?.TotalHaber ?? 0,
            FechaRegistro = proceso?.FechaRegistro,
            UsuarioRegistro = proceso?.UsuarioRegistro,
            Detalles = proceso?.Detalles
                .OrderBy(x => x.CodigoCuenta)
                .Select(x => new DiferenciaCambioProcesoDetalleItemViewModel
                {
                    IdPlanCuenta = x.IdPlanCuenta,
                    CodigoCuenta = x.CodigoCuenta,
                    NombreCuenta = x.NombreCuenta,
                    GeneraPorAnalisis = x.GeneraPorAnalisis,
                    TipoCambioAplicado = x.TipoCambioAplicado,
                    IdAsiento = x.IdAsiento,
                    NumeroAsiento = x.NumeroAsiento,
                    TotalDebe = x.TotalDebe,
                    TotalHaber = x.TotalHaber,
                    Estado = x.Estado,
                    Observacion = x.Observacion
                })
                .ToList() ?? [],
            AniosDisponibles = Enumerable.Range(anio - 5, 11).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };
    }

    private ProcesoAjusteCuentaViewModel ConstruirViewModelAjusteCuenta(
        short anio,
        byte mes,
        ConfiguracionContableProvisionDto? provisionAjuste,
        AjusteCuentaProcesoDto? proceso,
        bool usaTipoCambioSbs,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta)
    {
        return new ProcesoAjusteCuentaViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId!.Value,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            PeriodoConsulta = $"{anio:0000}{mes:00}",
            AnioSeleccionado = anio,
            MesSeleccionado = mes,
            IdOrigenConfigurado = provisionAjuste?.IdOrigen,
            OrigenConfiguradoTexto = provisionAjuste is null
                ? "No configurado"
                : $"{provisionAjuste.CodigoOrigen} - {provisionAjuste.NombreOrigen}",
            PuedeGenerar = provisionAjuste?.IdOrigen is not null,
            ProcesoGenerado = proceso is not null,
            FechaAsiento = proceso?.FechaAsiento,
            UsaTipoCambioSbs = usaTipoCambioSbs,
            TipoCambioCompra = tipoCambioCompra,
            TipoCambioVenta = tipoCambioVenta,
            TotalCuentas = proceso?.TotalCuentas ?? 0,
            TotalAsientos = proceso?.TotalAsientos ?? 0,
            TotalDebe = proceso?.TotalDebe ?? 0,
            TotalHaber = proceso?.TotalHaber ?? 0,
            FechaRegistro = proceso?.FechaRegistro,
            UsuarioRegistro = proceso?.UsuarioRegistro,
            Detalles = proceso?.Detalles
                .OrderBy(x => x.CodigoCuenta)
                .Select(x => new AjusteCuentaProcesoDetalleItemViewModel
                {
                    IdPlanCuenta = x.IdPlanCuenta,
                    CodigoCuenta = x.CodigoCuenta,
                    NombreCuenta = x.NombreCuenta,
                    CodigoMoneda = x.CodigoMoneda,
                    TipoCambioAplicado = x.TipoCambioAplicado,
                    TotalAnalisis = x.TotalAnalisis,
                    IdAsiento = x.IdAsiento,
                    NumeroAsiento = x.NumeroAsiento,
                    TotalDebe = x.TotalDebe,
                    TotalHaber = x.TotalHaber,
                    Estado = x.Estado,
                    Observacion = x.Observacion
                })
                .ToList() ?? [],
            AniosDisponibles = Enumerable.Range(anio - 5, 11).ToList(),
            MesesDisponibles = ListarMesesCalendario()
        };
    }

    private ProcesoAperturaViewModel ConstruirViewModelApertura(
        short anioApertura,
        short anioSaldo,
        byte mesSaldoHasta,
        ConfiguracionContableProvisionDto? provisionApertura,
        AperturaProcesoDto? proceso,
        bool usaTipoCambioSbs,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta)
    {
        var mesProceso = proceso?.MesSaldoHasta ?? mesSaldoHasta;

        return new ProcesoAperturaViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId!.Value,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioAperturaSeleccionado = anioApertura,
            AnioSaldo = proceso?.AnioSaldo ?? anioSaldo,
            MesSaldoHastaSeleccionado = mesProceso,
            PeriodoSaldoHasta = proceso?.PeriodoSaldoHasta ?? $"{anioSaldo:0000}{mesSaldoHasta:00}",
            IdOrigenConfigurado = provisionApertura?.IdOrigen,
            OrigenConfiguradoTexto = provisionApertura is null
                ? "No configurado"
                : $"{provisionApertura.CodigoOrigen} - {provisionApertura.NombreOrigen}",
            PuedeGenerar = provisionApertura?.IdOrigen is not null,
            ProcesoGenerado = proceso is not null,
            UsaTipoCambioSbs = proceso?.UsaTipoCambioSbs ?? usaTipoCambioSbs,
            TipoCambioCompra = proceso?.TipoCambioCompra ?? tipoCambioCompra,
            TipoCambioVenta = proceso?.TipoCambioVenta ?? tipoCambioVenta,
            FechaAsiento = proceso?.FechaAsiento,
            IdAsiento = proceso?.IdAsiento,
            NumeroAsiento = proceso?.NumeroAsiento,
            TotalLineas = proceso?.TotalLineas ?? 0,
            TotalDebe = proceso?.TotalDebe ?? 0,
            TotalHaber = proceso?.TotalHaber ?? 0,
            FechaRegistro = proceso?.FechaRegistro,
            UsuarioRegistro = proceso?.UsuarioRegistro,
            AniosDisponibles = Enumerable.Range(anioApertura - 5, 11).ToList(),
            MesesContablesDisponibles = ListarMesesContables(),
            Detalles = proceso?.Detalles
                .OrderBy(x => x.Item)
                .Select(x => new AperturaProcesoDetalleItemViewModel
                {
                    Item = x.Item,
                    TipoDetalle = x.TipoDetalle,
                    IdPlanCuenta = x.IdPlanCuenta,
                    CodigoCuenta = x.CodigoCuenta,
                    NombreCuenta = x.NombreCuenta,
                    CodigoMoneda = x.CodigoMoneda,
                    TipoCambioAplicado = x.TipoCambioAplicado,
                    TipoDocumento = x.TipoDocumento,
                    Serie = x.Serie,
                    NumeroDocumento = x.NumeroDocumento,
                    Debe = x.Debe,
                    Haber = x.Haber,
                    TotalImporteS = x.TotalImporteS,
                    TotalImporteD = x.TotalImporteD,
                    Observacion = x.Observacion
                })
                .ToList() ?? []
        };
    }

    private ProcesoCierreViewModel ConstruirViewModelCierre(
        short anio,
        ConfiguracionContableProvisionDto? provisionCierre,
        CierreProcesoDto? proceso,
        bool usaTipoCambioSbs,
        decimal tipoCambioCompra,
        decimal tipoCambioVenta)
    {
        return new ProcesoCierreViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId!.Value,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            AnioSeleccionado = anio,
            IdOrigenConfigurado = provisionCierre?.IdOrigen,
            OrigenConfiguradoTexto = provisionCierre is null
                ? "No configurado"
                : $"{provisionCierre.CodigoOrigen} - {provisionCierre.NombreOrigen}",
            PuedeGenerar = provisionCierre?.IdOrigen is not null,
            ProcesoGenerado = proceso is not null,
            UsaTipoCambioSbs = proceso?.UsaTipoCambioSbs ?? usaTipoCambioSbs,
            TipoCambioCompra = proceso?.TipoCambioCompra ?? tipoCambioCompra,
            TipoCambioVenta = proceso?.TipoCambioVenta ?? tipoCambioVenta,
            ProcesarGananciasPerdidas = proceso?.ProcesaGananciasPerdidas ?? true,
            ProcesarInventarios = proceso?.ProcesaInventarios ?? true,
            GananciasPerdidasGenerado = proceso?.Detalles.Any(x => x.TipoCierre == "14" && x.IdAsiento.HasValue) ?? false,
            InventariosGenerado = proceso?.Detalles.Any(x => x.TipoCierre == "15" && x.IdAsiento.HasValue) ?? false,
            FechaAsiento = proceso?.FechaAsiento,
            TotalCuentas = proceso?.TotalCuentas ?? 0,
            TotalAsientos = proceso?.TotalAsientos ?? 0,
            TotalDebe = proceso?.TotalDebe ?? 0,
            TotalHaber = proceso?.TotalHaber ?? 0,
            FechaRegistro = proceso?.FechaRegistro,
            UsuarioRegistro = proceso?.UsuarioRegistro,
            AniosDisponibles = Enumerable.Range(anio - 5, 11).ToList(),
            DetallesGananciasPerdidas = proceso?.Detalles
                .Where(x => x.TipoCierre == "14")
                .OrderBy(x => x.CodigoCuenta)
                .Select(MapearDetalleCierre)
                .ToList() ?? [],
            DetallesInventarios = proceso?.Detalles
                .Where(x => x.TipoCierre == "15")
                .OrderBy(x => x.CodigoCuenta)
                .Select(MapearDetalleCierre)
                .ToList() ?? []
        };
    }

    private static CierreProcesoDetalleItemViewModel MapearDetalleCierre(CierreProcesoDetalleDto detalle)
    {
        return new CierreProcesoDetalleItemViewModel
        {
            TipoCierre = detalle.TipoCierre,
            DescripcionCierre = detalle.DescripcionCierre,
            IdPlanCuenta = detalle.IdPlanCuenta,
            CodigoCuenta = detalle.CodigoCuenta,
            NombreCuenta = detalle.NombreCuenta,
            CodigoMoneda = detalle.CodigoMoneda,
            TipoCambioAplicado = detalle.TipoCambioAplicado,
            IdAsiento = detalle.IdAsiento,
            NumeroAsiento = detalle.NumeroAsiento,
            TotalDebe = detalle.TotalDebe,
            TotalHaber = detalle.TotalHaber,
            Estado = detalle.Estado,
            Observacion = detalle.Observacion
        };
    }
}
