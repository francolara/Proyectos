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
    IMonedaRepository monedaRepository,
    ITipoComprobanteRepository tipoComprobanteRepository) : Controller
{
    private const int TamanoPagina = 20;
    private const string CodigoDocumentoRucSunat = "6";

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
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return BadRequest(new { ok = false, mensaje = "Debe seleccionar una empresa activa." });
        }

        var tipoDocumento = (request.TipoDocumento ?? string.Empty).Trim();
        var numeroDocumento = (request.NumeroDocumento ?? string.Empty).Trim();
        var razonSocial = (request.RazonSocial ?? string.Empty).Trim();
        var nombres = (request.Nombres ?? string.Empty).Trim();
        var apellidos = (request.Apellidos ?? string.Empty).Trim();
        var telefono = string.IsNullOrWhiteSpace(request.Telefono) ? null : request.Telefono.Trim();
        var correo = string.IsNullOrWhiteSpace(request.Correo) ? null : request.Correo.Trim();
        var esJuridica = string.Equals(tipoDocumento, CodigoDocumentoRucSunat, StringComparison.OrdinalIgnoreCase);

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
            TipoPersona = esJuridica ? "J" : "N",
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
            ventas,
            null);

        model.Formulario = formulario;
        return model;
    }

    private static void NormalizarFormulario(VentaFormViewModel formulario)
    {
        formulario.Detalles = formulario.Detalles
            .Where(x => !string.IsNullOrWhiteSpace(x.Descripcion) || x.ImporteBruto > 0 || x.ValorUnitario > 0 || x.Cantidad > 0)
            .Select((x, index) =>
            {
                x.Item = (short)(index + 1);
                return x;
            })
            .ToList();
    }

    private void ValidarFormulario(VentaFormViewModel formulario)
    {
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

            totalDetalle += detalle.ImporteBruto;
        }

        if (formulario.ImporteTotal != formulario.BaseImponible + formulario.Igv + formulario.Isc + formulario.OtrosTributos + formulario.Redondeo)
        {
            ModelState.AddModelError(string.Empty, "El importe total debe ser igual a la suma de base imponible, IGV, ISC, otros tributos y redondeo.");
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
        IReadOnlyCollection<VentaResumenDto> ventas,
        VentaDto? ventaEditar)
    {
        var items = ventas
            .Select(x => new VentaResumenItemViewModel
            {
                IdVenta = x.IdVenta,
                NombreCliente = x.NombreCliente,
                EscenarioOperacion = x.EscenarioOperacion,
                FechaContabilizacion = x.FechaContabilizacion,
                Documento = $"{x.TipoComprobante} {x.Serie}-{x.Numero}",
                CodigoMoneda = x.CodigoMoneda,
                ImporteTotal = x.ImporteTotal,
                IdAsiento = x.IdAsiento,
                Estado = x.Estado
            })
            .ToList();

        var clienteSeleccionado = clientes.FirstOrDefault(x => x.IdCliente == (ventaEditar?.IdCliente ?? clientes.FirstOrDefault()?.IdCliente));

        return new VentaIndexViewModel
        {
            IdEmpresa = empresaId,
            EmpresaNombre = empresaNombre,
            PeriodoConsulta = periodo,
            AnioSeleccionado = anioSeleccionado,
            MesSeleccionado = mesSeleccionado,
            TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty,
            TotalVentas = items.Count,
            TotalImportePeriodo = items.Sum(x => x.ImporteTotal),
            AniosDisponibles = ConstruirAnios(anioSeleccionado),
            MesesDisponibles = ConstruirMeses(),
            Clientes = clientes.ToList(),
            ConfiguracionesVenta = configuraciones.ToList(),
            TiposDocumentoIdentidad = tiposDocumentoIdentidad.ToList(),
            Monedas = monedas.ToList(),
            TiposComprobante = tiposComprobante.ToList(),
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
                    TipoComprobante = tiposComprobante.FirstOrDefault()?.CodigoTipoComprobante ?? "01"
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
}
