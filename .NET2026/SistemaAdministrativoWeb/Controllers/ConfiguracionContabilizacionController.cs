using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Parametros;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CONFIGCONTABLE")]
public class ConfiguracionContabilizacionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IConfiguracionContabilizacionRepository configuracionRepository,
    IOrigenRepository origenRepository,
    IPlanCuentaRepository planCuentaRepository,
    IParametroEmpresaRepository parametroEmpresaRepository) : Controller
{
    private static readonly (string Modulo, string Titulo, string Resumen, string Descripcion, string Icono, string SufijoHtml)[] DefinicionesProvision =
    [
        ("COM", "Compras", "Origen y asiento automatico", "Define el origen contable y estado para generar asientos automaticos de compras.", "bi-cart-check", "compras"),
        ("VEN", "Ventas", "Origen y asiento automatico", "Define el origen contable y estado para generar asientos automaticos de ventas.", "bi-cash-stack", "ventas"),
        ("EGR", "Egresos", "Origen y asiento automatico", "Define el origen contable base para futuros movimientos operativos de egresos.", "bi-box-arrow-right", "egresos"),
        ("ING", "Ingresos", "Origen y asiento automatico", "Define el origen contable base para futuros movimientos operativos de ingresos.", "bi-box-arrow-in-left", "ingresos"),
        ("APNC", "Aplicaciones", "Origen y asiento automatico", "Define el origen contable base para futuras aplicaciones de notas de credito.", "bi-arrow-repeat", "aplicaciones"),
        ("DET", "Detracciones", "Origen y asiento automatico", "Define el origen contable del asiento adicional que aplica la 42 contra la cuenta SPOT en compras.", "bi-bank", "detracciones"),
        ("PER", "Percepciones", "Origen y asiento automatico", "Define el origen contable del asiento adicional que registra la percepcion en compras contra la cuenta parametrizada.", "bi-wallet2", "percepciones"),
        ("DIF", "Diferencia en cambio", "Origen del proceso mensual", "Define el origen contable que usara el proceso web de diferencia en cambio para generar asientos separados por cuenta.", "bi-currency-exchange", "diferencia-cambio"),
        ("AJU", "Ajuste de cuentas", "Origen del proceso mensual", "Define el origen contable que usara el proceso web de ajuste de cuentas para generar asientos separados por cuenta analitica.", "bi-sliders", "ajuste-cuentas"),
        ("APR", "Asiento de apertura", "Origen del proceso anual", "Define el origen contable que usara el proceso web de apertura anual para generar el asiento del periodo 00 usando saldos del anio anterior.", "bi-journal-plus", "asiento-apertura"),
        ("CIE", "Asiento de cierre", "Origen del proceso anual", "Define el origen contable que usara el cierre anual para generar un unico asiento compuesto con las cuentas configuradas como Inventario en el plan contable.", "bi-journal-x", "asiento-cierre")
    ];

    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;
    private const int TamanoListaParametros = 200;

    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var model = await ConstruirPantallaAsync(cancellationToken);
        return View(model);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarProvision(ConfiguracionProvisionFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        if (!ModelState.IsValid || !formulario.IdOrigen.HasValue)
        {
            TempData["ConfiguracionContabilizacionError"] = "Seleccione el origen contable para la provision.";
            return RedirectToAction(nameof(Index));
        }

        await configuracionRepository.GuardarProvisionAsync(
            currentCompanyAccessor.EmpresaId.Value,
            formulario.ModuloOperacion.Trim().ToUpperInvariant(),
            formulario.IdOrigen.Value,
            formulario.GeneraAsientoAutomatico,
            formulario.Activo,
            User.Identity?.Name,
            cancellationToken);

        TempData["ConfiguracionContabilizacionOk"] = "Provision contable guardada correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarDocumento(int idTipoComprobante, int? idCuentaVentaSoles, int? idCuentaVentaDolares, int? idCuentaCompraSoles, int? idCuentaCompraDolares, bool activo = true, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await configuracionRepository.GuardarDocumentoAsync(currentCompanyAccessor.EmpresaId.Value, idTipoComprobante, idCuentaVentaSoles, idCuentaVentaDolares, idCuentaCompraSoles, idCuentaCompraDolares, activo, User.Identity?.Name, cancellationToken);
        TempData["ConfiguracionContabilizacionOk"] = "Documento configurado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarImpuesto(int idTipoImpuesto, int? idPlanCuenta, bool activo = true, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await configuracionRepository.GuardarImpuestoAsync(currentCompanyAccessor.EmpresaId.Value, idTipoImpuesto, idPlanCuenta, activo, User.Identity?.Name, cancellationToken);
        TempData["ConfiguracionContabilizacionOk"] = "Impuesto configurado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> GuardarParametro(int idParametroEmpresa, string? valorParametro, bool activo = true, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var idEmpresa = currentCompanyAccessor.EmpresaId.Value;
        var parametroActual = await parametroEmpresaRepository.ObtenerAsync(idEmpresa, idParametroEmpresa, cancellationToken);
        if (parametroActual is null || string.Equals(parametroActual.TipoParametro, "NA", StringComparison.OrdinalIgnoreCase))
        {
            TempData["ConfiguracionContabilizacionError"] = "El parametro indicado no existe o no esta disponible para configuracion contable.";
            return RedirectToAction(nameof(Index));
        }

        var valorNormalizado = (valorParametro ?? string.Empty).Trim().ToUpperInvariant();
        if (!string.IsNullOrWhiteSpace(valorNormalizado))
        {
            var cuentas = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, valorNormalizado, null, 1, 10, false, false, cancellationToken);
            var cuenta = cuentas.Items.FirstOrDefault(x =>
                string.Equals(x.CodigoCuenta, valorNormalizado, StringComparison.OrdinalIgnoreCase)
                && x.Estado
                && x.AceptaMovimiento
                && x.NivelCuenta == 5);

            if (cuenta is null)
            {
                TempData["ConfiguracionContabilizacionError"] = "Seleccione una cuenta contable activa de nivel 5 para el parametro.";
                return RedirectToAction(nameof(Index));
            }
        }

        await parametroEmpresaRepository.GuardarAsync(new GuardarParametroEmpresaRequest
        {
            IdParametroEmpresa = parametroActual.IdParametroEmpresa,
            IdEmpresa = parametroActual.IdEmpresa,
            TipoParametro = parametroActual.TipoParametro,
            CodigoParametro = parametroActual.CodigoParametro,
            ValorParametro = valorNormalizado,
            DescripcionParametro = parametroActual.DescripcionParametro,
            FecIni = parametroActual.FecIni,
            FecFin = parametroActual.FecFin,
            Activo = activo,
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["ConfiguracionContabilizacionOk"] = "Parametro contable guardado correctamente.";
        return RedirectToAction(nameof(Index));
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idConfiguracionContabilizacion, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idConfiguracionContabilizacion, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(ConfiguracionContabilizacionFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);
        ValidarFormulario(formulario);

        if (!ModelState.IsValid)
        {
            var empresaIdError = currentCompanyAccessor.EmpresaId.Value;
            var configuracionesError = await configuracionRepository.ListarPorEmpresaAsync(empresaIdError, cancellationToken);
            var origenesError = await origenRepository.ListarPorEmpresaAsync(empresaIdError, false, cancellationToken);
            var cuentasError = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaIdError, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
            var modelError = ConstruirViewModel(configuracionesError, origenesError, cuentasError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }

        try
        {
            await configuracionRepository.GuardarAsync(new GuardarConfiguracionContabilizacionRequest
            {
                IdConfiguracionContabilizacion = formulario.IdConfiguracionContabilizacion,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                ModuloOperacion = formulario.ModuloOperacion.Trim().ToUpperInvariant(),
                EscenarioOperacion = formulario.EscenarioOperacion.Trim().ToUpperInvariant(),
                IdOrigen = formulario.IdOrigen!.Value,
                Descripcion = formulario.Descripcion.Trim(),
                GeneraAsientoAutomatico = formulario.GeneraAsientoAutomatico,
                UsaTipoCambio = formulario.UsaTipoCambio,
                Activo = formulario.Activo,
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarConfiguracionContabilizacionDetalleRequest
                    {
                        Orden = x.Orden,
                        ComponenteContable = x.ComponenteContable.Trim().ToUpperInvariant(),
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        NaturalezaMovimiento = x.NaturalezaMovimiento.Trim().ToUpperInvariant(),
                        Activo = x.Activo
                    })
                    .ToList()
            }, cancellationToken);

            TempData["ConfiguracionContabilizacionOk"] = formulario.IdConfiguracionContabilizacion.HasValue
                ? "Configuracion contable actualizada correctamente."
                : "Configuracion contable registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var empresaIdError = currentCompanyAccessor.EmpresaId.Value;
            var configuracionesError = await configuracionRepository.ListarPorEmpresaAsync(empresaIdError, cancellationToken);
            var origenesError = await origenRepository.ListarPorEmpresaAsync(empresaIdError, false, cancellationToken);
            var cuentasError = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaIdError, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
            var modelError = ConstruirViewModel(configuracionesError, origenesError, cuentasError, null);
            modelError.Formulario = formulario;
            return View("Formulario", modelError);
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idConfiguracionContabilizacion, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        await configuracionRepository.EliminarAsync(idConfiguracionContabilizacion, cancellationToken);
        TempData["ConfiguracionContabilizacionOk"] = "Configuracion contable eliminada correctamente.";
        return RedirectToAction(nameof(Index));
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idConfiguracionContabilizacion, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var configuraciones = await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken);
        var origenes = await origenRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken);
        var cuentas = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items.ToList();
        var configuracionEditar = idConfiguracionContabilizacion.HasValue
            ? await configuracionRepository.ObtenerAsync(idConfiguracionContabilizacion.Value, cancellationToken)
            : null;

        if (configuracionEditar is not null && configuracionEditar.IdEmpresa != empresaId)
        {
            configuracionEditar = null;
        }

        return View("Formulario", ConstruirViewModel(configuraciones, origenes, cuentas, configuracionEditar));
    }

    private async Task<ConfiguracionContabilizacionIndexViewModel> ConstruirPantallaAsync(CancellationToken cancellationToken)
    {
        var empresaId = currentCompanyAccessor.EmpresaId!.Value;

        var origenes = (await origenRepository.ListarPorEmpresaAsync(empresaId, false, cancellationToken))
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoOrigen)
            .ToList();
        var cuentas = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, false, false, cancellationToken)).Items
            .Where(x => x.Estado)
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var configuracion = await configuracionRepository.ObtenerConfiguracionContableEmpresaAsync(empresaId, cancellationToken);
        var parametros = (await parametroEmpresaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoListaParametros, cancellationToken)).Items
            .Where(x => !string.Equals(x.TipoParametro, "NA", StringComparison.OrdinalIgnoreCase))
            .OrderBy(x => x.TipoParametro)
            .ThenBy(x => x.CodigoParametro)
            .ToList();

        return new ConfiguracionContabilizacionIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            Origenes = origenes,
            CuentasMovimiento = cuentas,
            Provisiones = DefinicionesProvision
                .Select(definicion =>
                {
                    var provision = configuracion.Provisiones.FirstOrDefault(x => x.ModuloOperacion == definicion.Modulo);
                    return new ConfiguracionProvisionOperacionViewModel
                    {
                        ModuloOperacion = definicion.Modulo,
                        Titulo = definicion.Titulo,
                        Resumen = definicion.Resumen,
                        Descripcion = definicion.Descripcion,
                        Icono = definicion.Icono,
                        SufijoHtml = definicion.SufijoHtml,
                        Formulario = MapearProvision(definicion.Modulo, provision, origenes)
                    };
                })
                .ToList(),
            Documentos = configuracion.Documentos.Select(x => new ConfiguracionDocumentoFormViewModel
            {
                IdTipoComprobante = x.IdTipoComprobante,
                CodigoTipoComprobante = x.CodigoTipoComprobante,
                Descripcion = x.Descripcion,
                UsoCompras = x.UsoCompras,
                UsoVentas = x.UsoVentas,
                IdCuentaVentaSoles = x.IdCuentaVentaSoles,
                CuentaVentaSolesTexto = x.CuentaVentaSolesTexto,
                IdCuentaVentaDolares = x.IdCuentaVentaDolares,
                CuentaVentaDolaresTexto = x.CuentaVentaDolaresTexto,
                IdCuentaCompraSoles = x.IdCuentaCompraSoles,
                CuentaCompraSolesTexto = x.CuentaCompraSolesTexto,
                IdCuentaCompraDolares = x.IdCuentaCompraDolares,
                CuentaCompraDolaresTexto = x.CuentaCompraDolaresTexto,
                Activo = x.Activo
            }).ToList(),
            Impuestos = configuracion.Impuestos
                .Where(x => !string.Equals(x.CodigoSunat, "SPOT", StringComparison.OrdinalIgnoreCase))
                .Select(MapearImpuesto)
                .ToList(),
            Parametros = await MapearParametrosAsync(empresaId, parametros, cancellationToken)
        };
    }

    private static ConfiguracionProvisionFormViewModel MapearProvision(string moduloOperacion, ConfiguracionContableProvisionDto? provision, IReadOnlyCollection<OrigenDto> origenes)
    {
        var origenInicial = provision?.IdOrigen is null
                ? (string.Equals(moduloOperacion, "DIF", StringComparison.OrdinalIgnoreCase)
                ? origenes.FirstOrDefault(x => string.Equals(x.CodigoOrigen, "88", StringComparison.OrdinalIgnoreCase)) ?? origenes.FirstOrDefault()
                : string.Equals(moduloOperacion, "AJU", StringComparison.OrdinalIgnoreCase)
                    ? origenes.FirstOrDefault(x => string.Equals(x.CodigoOrigen, "67", StringComparison.OrdinalIgnoreCase)) ?? origenes.FirstOrDefault()
                    : string.Equals(moduloOperacion, "APR", StringComparison.OrdinalIgnoreCase)
                        ? origenes.FirstOrDefault(x => string.Equals(x.CodigoOrigen, "00", StringComparison.OrdinalIgnoreCase)) ?? origenes.FirstOrDefault()
                    : string.Equals(moduloOperacion, "CIE", StringComparison.OrdinalIgnoreCase)
                        ? origenes.FirstOrDefault(x => string.Equals(x.CodigoOrigen, "77", StringComparison.OrdinalIgnoreCase)) ?? origenes.FirstOrDefault()
                    : origenes.FirstOrDefault())
            : origenes.FirstOrDefault(x => x.IdOrigen == provision.IdOrigen.Value);

        return new ConfiguracionProvisionFormViewModel
        {
            ModuloOperacion = moduloOperacion,
            IdOrigen = provision?.IdOrigen ?? origenInicial?.IdOrigen,
            OrigenTexto = origenInicial is null ? "Seleccione origen" : $"{origenInicial.CodigoOrigen} - {origenInicial.NombreOrigen}",
            GeneraAsientoAutomatico = provision?.GeneraAsientoAutomatico ?? true,
            Activo = provision?.Activo ?? true
        };
    }

    private static ConfiguracionImpuestoFormViewModel MapearImpuesto(ConfiguracionImpuestoEmpresaDto impuesto)
    {
        return new ConfiguracionImpuestoFormViewModel
        {
            IdTipoImpuesto = impuesto.IdTipoImpuesto,
            CodigoSunat = impuesto.CodigoSunat,
            NombreImpuesto = impuesto.NombreImpuesto,
            IdPlanCuenta = impuesto.IdPlanCuenta,
            CuentaTexto = impuesto.CuentaTexto,
            Activo = impuesto.Activo
        };
    }

    private async Task<List<ConfiguracionParametroContableFormViewModel>> MapearParametrosAsync(int idEmpresa, IReadOnlyCollection<ParametroEmpresaDto> parametros, CancellationToken cancellationToken)
    {
        var items = new List<ConfiguracionParametroContableFormViewModel>();

        foreach (var parametro in parametros)
        {
            var cuentaTexto = string.Empty;
            if (!string.IsNullOrWhiteSpace(parametro.ValorParametro))
            {
                var cuentas = await planCuentaRepository.ListarPaginadoPorEmpresaAsync(idEmpresa, parametro.ValorParametro, null, 1, 10, false, false, cancellationToken);
                var cuenta = cuentas.Items.FirstOrDefault(x => string.Equals(x.CodigoCuenta, parametro.ValorParametro, StringComparison.OrdinalIgnoreCase));
                cuentaTexto = cuenta is null ? parametro.ValorParametro : $"{cuenta.CodigoCuenta} - {cuenta.NombreCuenta}";
            }

            items.Add(new ConfiguracionParametroContableFormViewModel
            {
                IdParametroEmpresa = parametro.IdParametroEmpresa,
                TipoParametro = parametro.TipoParametro,
                CodigoParametro = parametro.CodigoParametro,
                DescripcionParametro = parametro.DescripcionParametro,
                ValorParametro = parametro.ValorParametro,
                CuentaTexto = cuentaTexto,
                Activo = parametro.Activo
            });
        }

        return items;
    }

    private static void NormalizarFormulario(ConfiguracionContabilizacionFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => x.IdPlanCuenta.HasValue || !string.IsNullOrWhiteSpace(x.ComponenteContable))
            .Select((x, index) =>
            {
                x.Orden = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(ConfiguracionContabilizacionFormViewModel formulario)
    {
        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos un componente contable.");
            return;
        }

        var componentesActivos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (!detalle.IdPlanCuenta.HasValue)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione una cuenta.");
            }

            if (detalle.Activo && !componentesActivos.Add(detalle.ComponenteContable))
            {
                ModelState.AddModelError($"{prefijo}.ComponenteContable", "No repita el mismo componente activo.");
            }
        }
    }

    private ConfiguracionContabilizacionIndexViewModel ConstruirViewModel(
        IReadOnlyCollection<ConfiguracionContabilizacionResumenDto> configuraciones,
        IReadOnlyCollection<OrigenDto> origenes,
        IReadOnlyCollection<PlanCuentaDto> cuentas,
        ConfiguracionContabilizacionDto? configuracionEditar)
    {
        var items = configuraciones
            .Select(x => new ConfiguracionContabilizacionResumenItemViewModel
            {
                IdConfiguracionContabilizacion = x.IdConfiguracionContabilizacion,
                ModuloOperacion = x.ModuloOperacion,
                EscenarioOperacion = x.EscenarioOperacion,
                CodigoOrigen = x.CodigoOrigen,
                NombreOrigen = x.NombreOrigen,
                Descripcion = x.Descripcion,
                GeneraAsientoAutomatico = x.GeneraAsientoAutomatico,
                UsaTipoCambio = x.UsaTipoCambio,
                Activo = x.Activo,
                CantidadComponentes = x.CantidadComponentes
            })
            .OrderBy(x => x.ModuloOperacion)
            .ThenBy(x => x.EscenarioOperacion)
            .ToList();

        return new ConfiguracionContabilizacionIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TotalConfiguraciones = items.Count,
            TotalActivas = items.Count(x => x.Activo),
            TotalAutomaticas = items.Count(x => x.GeneraAsientoAutomatico),
            Configuraciones = items,
            Origenes = origenes.Where(x => x.Estado).OrderBy(x => x.CodigoOrigen).ToList(),
            CuentasMovimiento = cuentas.Where(x => x.Estado).OrderBy(x => x.CodigoCuenta).ToList(),
            Formulario = configuracionEditar is null
                ? new ConfiguracionContabilizacionFormViewModel
                {
                    IdOrigen = origenes.FirstOrDefault(x => x.Estado)?.IdOrigen
                }
                : new ConfiguracionContabilizacionFormViewModel
                {
                    IdConfiguracionContabilizacion = configuracionEditar.IdConfiguracionContabilizacion,
                    ModuloOperacion = configuracionEditar.ModuloOperacion,
                    EscenarioOperacion = configuracionEditar.EscenarioOperacion,
                    IdOrigen = configuracionEditar.IdOrigen,
                    Descripcion = configuracionEditar.Descripcion,
                    GeneraAsientoAutomatico = configuracionEditar.GeneraAsientoAutomatico,
                    UsaTipoCambio = configuracionEditar.UsaTipoCambio,
                    Activo = configuracionEditar.Activo,
                    Detalles = configuracionEditar.Detalles
                        .OrderBy(x => x.Orden)
                        .Select(x => new ConfiguracionContabilizacionDetalleFormViewModel
                        {
                            Orden = x.Orden,
                            ComponenteContable = x.ComponenteContable,
                            IdPlanCuenta = x.IdPlanCuenta,
                            CuentaTexto = $"{x.CodigoCuenta} - {x.NombreCuenta}",
                            NaturalezaMovimiento = x.NaturalezaMovimiento,
                            Activo = x.Activo
                        })
                        .ToList()
                }
        };
    }
}
