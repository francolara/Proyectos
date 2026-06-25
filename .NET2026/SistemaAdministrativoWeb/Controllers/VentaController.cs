using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class VentaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IVentaRepository ventaRepository,
    IClienteRepository clienteRepository,
    IPersonaRepository personaRepository,
    IConfiguracionContabilizacionRepository configuracionRepository,
    IAsientoPreviewService asientoPreviewService,
    IPlanCuentaRepository planCuentaRepository,
    ITipoAfectacionIgvRepository tipoAfectacionIgvRepository,
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const int TamanoAyudaCuenta = 100;
    private const string CodigoDocumentoRucSunat = "6";
    private const string CodigoAfectacionGravadoOnerosa = "10";
    private const decimal TasaIgv = 0.18m;

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
        var periodoTrabajo = $"{anioTrabajo:0000}{mesTrabajo:00}";
        var clientes = (await clienteRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "VEN")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(false, true, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var ventas = await ventaRepository.ListarPaginadoPorEmpresaAsync(empresaId, anioTrabajo, mesTrabajo, textoBusqueda, pagina, TamanoPagina, cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            anioTrabajo,
            mesTrabajo,
            textoBusqueda,
            clientes,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            ventas.Items,
            null);
        model.TotalVentas = ventas.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = ventas.TotalRecords
        };
        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idVenta, string? periodo = null, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(periodo, idVenta, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Eliminar(int idVenta, short? anio = null, byte? mes = null, string? textoBusqueda = null, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var (anioTrabajo, mesTrabajo) = NormalizarPeriodo(anio, mes);

        try
        {
            await ventaRepository.EliminarAsync(idVenta, currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            TempData["VentaOk"] = "Venta eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["VentaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { anio = anioTrabajo, mes = mesTrabajo, textoBusqueda, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(VentaFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        NormalizarFormulario(formulario);
        ValidarFormulario(formulario);

        var periodoTrabajo = $"{formulario.FechaContabilizacion.Year:0000}{formulario.FechaContabilizacion.Month:00}";

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            var result = await ventaRepository.GuardarAsync(new GuardarVentaRequest
            {
                IdVenta = formulario.IdVenta,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                IdCliente = formulario.IdCliente!.Value,
                IdConfiguracionContabilizacion = formulario.IdConfiguracionContabilizacion!.Value,
                FechaEmision = formulario.FechaEmision,
                FechaContabilizacion = formulario.FechaContabilizacion,
                TipoComprobante = formulario.TipoComprobante.Trim().ToUpperInvariant(),
                Serie = formulario.Serie.Trim().ToUpperInvariant(),
                Numero = formulario.Numero.Trim().ToUpperInvariant(),
                IdMoneda = formulario.IdMoneda!.Value,
                TipoCambio = formulario.TipoCambio,
                BaseImponible = formulario.BaseImponible,
                TotalExonerado = formulario.TotalExonerado,
                TotalInafecto = formulario.TotalInafecto,
                Icbper = formulario.Icbper,
                Igv = formulario.Igv,
                Isc = formulario.Isc,
                OtrosTributos = formulario.OtrosTributos,
                Redondeo = formulario.Redondeo,
                ImporteTotal = formulario.ImporteTotal,
                Observacion = string.IsNullOrWhiteSpace(formulario.Observacion) ? null : formulario.Observacion.Trim(),
                UsuarioRegistro = User.Identity?.Name,
                Detalles = formulario.Detalles
                    .Select(x => new GuardarVentaDetalleRequest
                    {
                        Item = x.Item,
                        IdPlanCuenta = x.IdPlanCuenta!.Value,
                        IdTipoAfectacionIGV = x.IdTipoAfectacionIGV!.Value,
                        Descripcion = x.Descripcion.Trim(),
                        Cantidad = x.Cantidad,
                        ValorUnitario = x.ValorUnitario,
                        ImporteBruto = x.ImporteBruto
                    })
                    .ToList()
            }, cancellationToken);

            TempData["VentaOk"] = $"Venta registrada correctamente. Asiento vinculado: {(result.IdAsiento.HasValue ? result.IdAsiento.Value.ToString() : "sin asiento")}.";
            return RedirectToAction(nameof(Index), new { anio = formulario.FechaContabilizacion.Year, mes = formulario.FechaContabilizacion.Month });
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirViewModelErrorAsync(currentCompanyAccessor.EmpresaId.Value, periodoTrabajo, formulario, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    [HttpGet]
    public async Task<IActionResult> BuscarClientes(string? buscar = null, int? clienteId = null, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return Json(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var criterio = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var clientes = await clienteRepository.ListarActivosPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);

        var items = clientes
            .Where(x =>
                clienteId.HasValue && x.IdCliente == clienteId.Value
                || criterio is null
                || x.NombreCompleto.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || x.NumeroDocumento.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || x.CodigoCliente.Contains(criterio, StringComparison.OrdinalIgnoreCase)
                || (!string.IsNullOrWhiteSpace(x.Telefono) && x.Telefono.Contains(criterio, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(x.CorreoElectronico) && x.CorreoElectronico.Contains(criterio, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(x => x.NombreCompleto)
            .Take(30)
            .Select(x => new
            {
                value = x.IdCliente,
                text = $"{x.NombreCompleto} ({x.NumeroDocumento})",
                tipoDocumento = x.TipoDocumento,
                numeroDocumento = x.NumeroDocumento,
                nombre = x.NombreCompleto,
                numero = x.Telefono ?? string.Empty,
                correo = x.CorreoElectronico ?? string.Empty
            })
            .ToList();

        return Json(new { ok = true, items });
    }

    [HttpPost]
    public async Task<IActionResult> CrearClienteRapido([FromBody] RegistroRapidoPersonaRequestViewModel request, CancellationToken cancellationToken = default)
    {
        try
        {
            if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
            {
                return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
            }

            var tipoPersona = (request.TipoPersona ?? string.Empty).Trim().ToUpperInvariant();
            var tipoDocumento = (request.TipoDocumento ?? string.Empty).Trim();
            var numeroDocumento = (request.NumeroDocumento ?? string.Empty).Trim();
            var razonSocial = (request.RazonSocial ?? string.Empty).Trim();
            var nombres = (request.Nombres ?? string.Empty).Trim();
            var apellidos = (request.Apellidos ?? string.Empty).Trim();
            var telefono = string.IsNullOrWhiteSpace(request.Telefono) ? null : request.Telefono.Trim();
            var correo = string.IsNullOrWhiteSpace(request.Correo) ? null : request.Correo.Trim();
            var esJuridica = string.Equals(tipoPersona, "J", StringComparison.OrdinalIgnoreCase);

            if (!string.Equals(tipoPersona, "N", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(tipoPersona, "J", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(new { ok = false, mensaje = "Seleccione el tipo de persona." });
            }

            if (string.IsNullOrWhiteSpace(tipoDocumento))
            {
                return BadRequest(new { ok = false, mensaje = "Seleccione el tipo de documento." });
            }

            if (string.IsNullOrWhiteSpace(numeroDocumento))
            {
                return BadRequest(new { ok = false, mensaje = "Ingrese el numero de documento." });
            }

            if (esJuridica)
            {
                if (string.IsNullOrWhiteSpace(razonSocial))
                {
                    return BadRequest(new { ok = false, mensaje = "Ingrese la razon social del cliente." });
                }
            }
            else if (string.IsNullOrWhiteSpace(nombres) || string.IsNullOrWhiteSpace(apellidos))
            {
                return BadRequest(new { ok = false, mensaje = "Ingrese nombres y apellidos del cliente." });
            }

            await personaRepository.GuardarAsync(new GuardarPersonaRequest
            {
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                TipoPersona = tipoPersona,
                TipoDocumento = tipoDocumento,
                NumeroDocumento = numeroDocumento,
                ApellidoPaterno = esJuridica ? null : apellidos,
                ApellidoMaterno = null,
                Nombres = esJuridica ? null : nombres,
                RazonSocial = esJuridica ? razonSocial : null,
                CorreoElectronico = correo,
                Telefono = telefono,
                Direccion = null,
                CodigoUbigeo = null,
                EsCliente = true,
                EsProveedor = false,
                Estado = true,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            var clientes = await clienteRepository.ListarActivosPorEmpresaAsync(currentCompanyAccessor.EmpresaId.Value, cancellationToken);
            var cliente = clientes.FirstOrDefault(x =>
                string.Equals(x.TipoDocumento, tipoDocumento, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.NumeroDocumento, numeroDocumento, StringComparison.OrdinalIgnoreCase));

            if (cliente is null)
            {
                return BadRequest(new { ok = false, mensaje = "El cliente fue registrado, pero no pudo recuperarse para la seleccion." });
            }

            return Json(new
            {
                ok = true,
                clienteId = cliente.IdCliente,
                clienteTexto = $"{cliente.NombreCompleto} ({cliente.NumeroDocumento})",
                tipoDocumento = cliente.TipoDocumento,
                numeroDocumento = cliente.NumeroDocumento,
                nombre = cliente.NombreCompleto,
                numero = cliente.Telefono ?? string.Empty,
                correo = cliente.CorreoElectronico ?? string.Empty
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> PrevisualizarAsiento([FromBody] AsientoPreviewRequestViewModel request, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        if (!request.IdConfiguracionContabilizacion.HasValue || request.IdConfiguracionContabilizacion.Value <= 0)
        {
            return BadRequest(new { ok = false, mensaje = "Seleccione una configuracion contable." });
        }

        try
        {
            var preview = await asientoPreviewService.PrevisualizarAsync(currentCompanyAccessor.EmpresaId.Value, new AsientoPreviewRequest
            {
                ModuloOperacion = "VEN",
                IdConfiguracionContabilizacion = request.IdConfiguracionContabilizacion.Value,
                FechaContabilizacion = request.FechaContabilizacion,
                TipoComprobante = request.TipoComprobante ?? string.Empty,
                Serie = request.Serie ?? string.Empty,
                Numero = request.Numero ?? string.Empty,
                BaseImponible = request.BaseImponible,
                Igv = request.Igv,
                Isc = request.Isc,
                OtrosTributos = request.OtrosTributos,
                Redondeo = request.Redondeo,
                ImporteTotal = request.ImporteTotal,
                Detalles = request.Detalles
                    .Select(x => new AsientoPreviewDetalleRequest
                    {
                        Item = x.Item,
                        Descripcion = x.Descripcion,
                        Cantidad = x.Cantidad,
                        ValorUnitario = x.ValorUnitario,
                        ImporteBruto = x.ImporteBruto
                    })
                    .ToList()
            }, cancellationToken);

            return Json(new { ok = true, preview });
        }
        catch (Exception ex)
        {
            return BadRequest(new { ok = false, mensaje = ex.Message });
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(string? periodo, int? idVenta, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var empresaId = currentCompanyAccessor.EmpresaId.Value;
        var periodoTrabajo = NormalizarPeriodo(periodo);
        var clientes = (await clienteRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "VEN")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(false, true, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var ventas = await ventaRepository.ListarPorEmpresaAsync(empresaId, periodoTrabajo, cancellationToken);
        var ventaEditar = idVenta.HasValue
            ? await ventaRepository.ObtenerAsync(idVenta.Value, cancellationToken)
            : null;

        if (ventaEditar is not null && ventaEditar.IdEmpresa != empresaId)
        {
            ventaEditar = null;
        }

        return View("Formulario", ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodoTrabajo,
            short.Parse(periodoTrabajo[..4]),
            byte.Parse(periodoTrabajo[4..]),
            null,
            clientes,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            ventas,
            ventaEditar));
    }

    private async Task<VentaIndexViewModel> ConstruirViewModelErrorAsync(int empresaId, string periodo, VentaFormViewModel formulario, CancellationToken cancellationToken)
    {
        var clientes = (await clienteRepository.ListarActivosPorEmpresaAsync(empresaId, cancellationToken))
            .OrderBy(x => x.NombreCompleto)
            .ToList();
        var configuraciones = (await configuracionRepository.ListarPorEmpresaAsync(empresaId, cancellationToken))
            .Where(x => x.Activo && x.ModuloOperacion == "VEN")
            .OrderBy(x => x.EscenarioOperacion)
            .ToList();
        var tiposDocumentoIdentidad = (await personaRepository.ListarTiposDocumentoAsync(cancellationToken))
            .OrderBy(x => x.Orden)
            .ThenBy(x => x.CodigoSunat)
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoSunat,
                Texto = $"{x.CodigoSunat} - {x.Nombre}"
            })
            .ToList();
        var monedas = (await monedaRepository.ListarActivasAsync(cancellationToken))
            .OrderByDescending(x => x.EsMonedaBase)
            .ThenBy(x => x.CodigoMoneda)
            .ToList();
        var tiposComprobante = (await tipoComprobanteRepository.ListarActivosAsync(false, true, cancellationToken))
            .OrderBy(x => x.CodigoTipoComprobante)
            .ToList();
        var cuentasMovimiento = (await planCuentaRepository.ListarPaginadoPorEmpresaAsync(empresaId, null, null, 1, TamanoAyudaCuenta, true, false, cancellationToken)).Items
            .OrderBy(x => x.CodigoCuenta)
            .ToList();
        var tiposAfectacionIgv = (await tipoAfectacionIgvRepository.ListarActivosAsync(cancellationToken))
            .OrderBy(x => x.CodigoSunat)
            .ToList();
        var ventas = await ventaRepository.ListarPorEmpresaAsync(empresaId, periodo, cancellationToken);

        var model = ConstruirViewModel(
            empresaId,
            currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            periodo,
            short.Parse(periodo[..4]),
            byte.Parse(periodo[4..]),
            null,
            clientes,
            configuraciones,
            tiposDocumentoIdentidad,
            monedas,
            tiposComprobante,
            cuentasMovimiento,
            tiposAfectacionIgv,
            ventas,
            null);

        model.Formulario = formulario;
        return model;
    }

    private static void NormalizarFormulario(VentaFormViewModel formulario)
    {
        formulario.Serie = NormalizarSerieDocumento(formulario.Serie);
        formulario.Numero = NormalizarNumeroDocumento(formulario.Numero);
        formulario.Detalles = formulario.Detalles
            .Where(x => !string.IsNullOrWhiteSpace(x.Descripcion) || x.ImporteBruto > 0 || x.ValorUnitario > 0 || x.Cantidad > 0 || x.IdPlanCuenta.HasValue)
            .Select((x, index) =>
            {
                x.Item = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(VentaFormViewModel formulario)
    {
        if (formulario.IdCliente.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdCliente), "Seleccione un cliente.");
        }

        if (formulario.IdConfiguracionContabilizacion.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdConfiguracionContabilizacion), "Seleccione una configuracion contable.");
        }

        if (formulario.IdMoneda.GetValueOrDefault() <= 0)
        {
            ModelState.AddModelError(nameof(formulario.IdMoneda), "Seleccione la moneda.");
        }

        if (string.IsNullOrWhiteSpace(formulario.TipoComprobante))
        {
            ModelState.AddModelError(nameof(formulario.TipoComprobante), "Seleccione el tipo de comprobante.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Serie))
        {
            ModelState.AddModelError(nameof(formulario.Serie), "Ingrese la serie del documento.");
        }

        if (string.IsNullOrWhiteSpace(formulario.Numero))
        {
            ModelState.AddModelError(nameof(formulario.Numero), "Ingrese el numero del documento.");
        }

        if (formulario.Detalles.Count == 0)
        {
            ModelState.AddModelError(string.Empty, "Debe registrar al menos un concepto en la venta.");
        }

        decimal totalDetalle = 0;

        for (var i = 0; i < formulario.Detalles.Count; i++)
        {
            var detalle = formulario.Detalles[i];
            var prefijo = $"Formulario.Detalles[{i}]";

            if (string.IsNullOrWhiteSpace(detalle.Descripcion))
            {
                ModelState.AddModelError($"{prefijo}.Descripcion", "Ingrese la descripcion del concepto.");
            }

            if (!detalle.IdPlanCuenta.HasValue || detalle.IdPlanCuenta.Value <= 0)
            {
                ModelState.AddModelError($"{prefijo}.IdPlanCuenta", "Seleccione la cuenta contable.");
            }

            if (!detalle.IdTipoAfectacionIGV.HasValue || detalle.IdTipoAfectacionIGV.Value <= 0)
            {
                ModelState.AddModelError($"{prefijo}.IdTipoAfectacionIGV", "Seleccione el tipo de afectacion IGV.");
            }

            totalDetalle += detalle.ImporteBruto;
        }

        if (formulario.ImporteTotal != formulario.BaseImponible + formulario.Igv)
        {
            ModelState.AddModelError(string.Empty, "El importe total debe ser igual a la suma del subtotal e IGV.");
        }

        if (formulario.BaseImponible > 0 && totalDetalle > 0 && decimal.Round(totalDetalle, 2) != decimal.Round(formulario.BaseImponible, 2))
        {
            ModelState.AddModelError(string.Empty, "La suma del detalle debe coincidir con la base imponible.");
        }
    }

    private static (short anio, byte mes) NormalizarPeriodo(short? anio, byte? mes)
    {
        var today = DateTime.Today;
        return (anio ?? (short)today.Year, mes is >= 1 and <= 12 ? mes.Value : (byte)today.Month);
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

    private static VentaIndexViewModel ConstruirViewModel(
        int empresaId,
        string empresaNombre,
        string periodo,
        short anioSeleccionado,
        byte mesSeleccionado,
        string? textoBusqueda,
        IReadOnlyCollection<ClienteDto> clientes,
        IReadOnlyCollection<ConfiguracionContabilizacionResumenDto> configuraciones,
        IReadOnlyCollection<OpcionCatalogoViewModel> tiposDocumentoIdentidad,
        IReadOnlyCollection<MonedaDto> monedas,
        IReadOnlyCollection<TipoComprobanteDto> tiposComprobante,
        IReadOnlyCollection<PlanCuentaDto> cuentasMovimiento,
        IReadOnlyCollection<TipoAfectacionIgvDto> tiposAfectacionIgv,
        IReadOnlyCollection<VentaResumenDto> ventas,
        VentaDto? ventaEditar)
    {
        var items = ventas
            .Select(x => new VentaResumenItemViewModel
            {
                IdVenta = x.IdVenta,
                NombreCliente = x.NombreCliente,
                EscenarioOperacion = x.EscenarioOperacion,
                FechaEmision = x.FechaEmision,
                FechaContabilizacion = x.FechaContabilizacion,
                Documento = $"{x.TipoComprobante} {x.Serie}-{x.Numero}",
                CodigoMoneda = x.CodigoMoneda,
                ImporteTotal = x.ImporteTotal,
                Saldo = x.Saldo,
                IdAsiento = x.IdAsiento,
                Estado = x.Estado,
                Situacion = x.Situacion
            })
            .ToList();

        var clienteSeleccionado = clientes.FirstOrDefault(x => x.IdCliente == (ventaEditar?.IdCliente ?? clientes.FirstOrDefault()?.IdCliente));
        var totalImportePeriodo = items.Sum(x => x.ImporteTotal);
        var totalImporteSolesPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "PEN", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.ImporteTotal);
        var totalImporteDolaresPeriodo = items
            .Where(x => string.Equals(x.CodigoMoneda, "USD", StringComparison.OrdinalIgnoreCase))
            .Sum(x => x.ImporteTotal);

        return new VentaIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TotalVentas = items.Count,
            TotalImportePeriodo = totalImportePeriodo,
            TotalImporteSolesPeriodo = totalImporteSolesPeriodo,
            TotalImporteDolaresPeriodo = totalImporteDolaresPeriodo,
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            Clientes = clientes.ToList(),
            ConfiguracionesVenta = configuraciones.ToList(),
            TiposDocumentoIdentidad = tiposDocumentoIdentidad.ToList(),
            Monedas = monedas.ToList(),
            TiposComprobante = tiposComprobante.ToList(),
            CuentasMovimiento = cuentasMovimiento.ToList(),
            TiposAfectacionIgv = tiposAfectacionIgv.ToList(),
            Ventas = items,
            ClienteSeleccionadoTipoDocumento = clienteSeleccionado?.TipoDocumento ?? ventaEditar?.TipoDocumentoCliente ?? string.Empty,
            ClienteSeleccionadoNumeroDocumento = clienteSeleccionado?.NumeroDocumento ?? ventaEditar?.NumeroDocumentoCliente ?? string.Empty,
            ClienteSeleccionadoNombreLegal = clienteSeleccionado?.NombreCompleto ?? string.Empty,
            ClienteSeleccionadoTexto = clienteSeleccionado is null ? string.Empty : $"{clienteSeleccionado.NombreCompleto} ({clienteSeleccionado.NumeroDocumento})",
            ClienteSeleccionadoTelefono = clienteSeleccionado?.Telefono ?? string.Empty,
            ClienteSeleccionadoCorreo = clienteSeleccionado?.CorreoElectronico ?? string.Empty,
            Formulario = ventaEditar is null
                ? new VentaFormViewModel
                {
                    FechaEmision = ParsePeriodo(periodo),
                    FechaContabilizacion = ParsePeriodo(periodo),
                    IdMoneda = monedas.OrderByDescending(x => x.EsMonedaBase).FirstOrDefault()?.IdMoneda,
                    IdCliente = clientes.FirstOrDefault()?.IdCliente,
                    IdConfiguracionContabilizacion = configuraciones.FirstOrDefault()?.IdConfiguracionContabilizacion,
                    TipoComprobante = tiposComprobante.FirstOrDefault()?.CodigoTipoComprobante ?? "01",
                    Detalles =
                    [
                        new()
                        {
                            Item = 1,
                            IdTipoAfectacionIGV = tiposAfectacionIgv.FirstOrDefault(x => x.CodigoSunat == CodigoAfectacionGravadoOnerosa)?.IdTipoAfectacionIGV ?? 1
                        }
                    ]
                }
                : new VentaFormViewModel
                {
                    IdVenta = ventaEditar.IdVenta,
                    IdCliente = ventaEditar.IdCliente,
                    IdConfiguracionContabilizacion = ventaEditar.IdConfiguracionContabilizacion,
                    FechaEmision = ventaEditar.FechaEmision,
                    FechaContabilizacion = ventaEditar.FechaContabilizacion,
                    TipoComprobante = ventaEditar.TipoComprobante,
                    Serie = ventaEditar.Serie,
                    Numero = ventaEditar.Numero,
                    IdMoneda = ventaEditar.IdMoneda,
                    TipoCambio = ventaEditar.TipoCambio,
                    BaseImponible = ventaEditar.BaseImponible,
                    TotalExonerado = ventaEditar.TotalExonerado,
                    TotalInafecto = ventaEditar.TotalInafecto,
                    Icbper = ventaEditar.Icbper,
                    Igv = ventaEditar.Igv,
                    Isc = ventaEditar.Isc,
                    OtrosTributos = ventaEditar.OtrosTributos,
                    Redondeo = ventaEditar.Redondeo,
                    ImporteTotal = ventaEditar.ImporteTotal,
                    Observacion = ventaEditar.Observacion,
                    Detalles = ventaEditar.Detalles
                        .OrderBy(x => x.Item)
                        .Select(x => new VentaDetalleFormViewModel
                        {
                            Item = x.Item,
                            IdPlanCuenta = x.IdPlanCuenta,
                            CuentaTexto = $"{x.CodigoCuenta} - {x.NombreCuenta}",
                            IdTipoAfectacionIGV = x.IdTipoAfectacionIGV,
                            Descripcion = x.Descripcion,
                            Cantidad = x.Cantidad,
                            ValorUnitario = x.ValorUnitario,
                            ImporteBruto = x.ImporteBruto
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

    private static string NormalizarSerieDocumento(string? serie)
    {
        var serieNormalizada = new string((serie ?? string.Empty)
            .Trim()
            .ToUpperInvariant()
            .Where(char.IsLetterOrDigit)
            .ToArray());

        if (string.IsNullOrEmpty(serieNormalizada))
        {
            return string.Empty;
        }

        var prefijo = serieNormalizada[0];
        if (prefijo is 'F' or 'B')
        {
            var digitos = new string(serieNormalizada.Skip(1).Where(char.IsDigit).ToArray());
            digitos = digitos.Length > 3 ? digitos[..3] : digitos.PadLeft(3, '0');
            return $"{prefijo}{digitos}";
        }

        return serieNormalizada.Length > 4 ? serieNormalizada[..4] : serieNormalizada;
    }

    private static string NormalizarNumeroDocumento(string? numero)
    {
        var digitos = new string((numero ?? string.Empty).Where(char.IsDigit).ToArray());
        if (string.IsNullOrEmpty(digitos))
        {
            return string.Empty;
        }

        return digitos.Length > 10 ? digitos[..10] : digitos;
    }
}
