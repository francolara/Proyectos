using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;
using System.Linq;
using System.Text.RegularExpressions;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class ClientesController(IModuloPermisoService moduloPermisoService, ISportCenterStoredProcedureService spService)
    : ModuloControllerBase(moduloPermisoService)
{
    private const string CodigoDocumentoRucSunat = "6";
    private const string CodigoDocumentoNoDomiciliadoSinRucSunat = "0";

    public async Task<IActionResult> Index(int? negocioId, string? estado = null, string? buscar = null, int pagina = 1)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje)) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var estadoNormalizado = (estado ?? "activos").Trim().ToLowerInvariant();
        bool? activoFiltro = estadoNormalizado switch
        {
            "activos" => true,
            "inactivos" => false,
            _ => null
        };

        if (estadoNormalizado is not ("todos" or "activos" or "inactivos"))
            estadoNormalizado = "activos";

        const int tamanoPagina = 20;
        var paginaActual = pagina < 1 ? 1 : pagina;
        var textoBusqueda = string.IsNullOrWhiteSpace(buscar) ? null : buscar.Trim();
        var (clientesPagina, totalRegistros, totalActivos, totalInactivos) = await spService.ClientesListarAsync(
            resolvedNegocioId.Value,
            activoFiltro,
            textoBusqueda,
            paginaActual,
            tamanoPagina);

        var totalPaginas = Math.Max(1, (int)Math.Ceiling(totalRegistros / (double)tamanoPagina));
        if (paginaActual > totalPaginas)
        {
            paginaActual = totalPaginas;
            (clientesPagina, totalRegistros, totalActivos, totalInactivos) = await spService.ClientesListarAsync(
                resolvedNegocioId.Value,
                activoFiltro,
                textoBusqueda,
                paginaActual,
                tamanoPagina);
        }

        var vm = new ClientesIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            ModuloCodigo = baseVm.ModuloCodigo,
            ModuloNombre = baseVm.ModuloNombre,
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            EstadoFiltro = estadoNormalizado,
            Buscar = textoBusqueda,
            Pagina = paginaActual,
            TamanoPagina = tamanoPagina,
            TotalRegistros = totalRegistros,
            TotalPaginas = totalPaginas,
            TotalActivos = totalActivos,
            TotalInactivos = totalInactivos,
            Clientes = clientesPagina
        };
        return View(vm);
    }

    public async Task<IActionResult> Create(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = new ClienteFormViewModel
        {
            NegocioId = resolvedNegocioId.Value,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            TipoDocumento = CodigoDocumentoNoDomiciliadoSinRucSunat,
            Activo = true,
            CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais("+51")
        };
        await CargarCombosClienteAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeCrear) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        NormalizarYValidarIdentidad(model);
        ComponerTelefono(model);
        await NormalizarYValidarUbigeoAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCombosClienteAsync(model);
            return View(model);
        }

        try
        {
            await spService.ClientesCrearAsync(model, User.Identity?.Name ?? "sistema");
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (SqlException ex) when (EsErrorClienteDuplicado(ex.Message))
        {
            ModelState.AddModelError(string.Empty, "Cliente ya se encuentra registrado.");
            await CargarCombosClienteAsync(model);
            return View(model);
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            await CargarCombosClienteAsync(model);
            return View(model);
        }
    }

    public async Task<IActionResult> Edit(int id, int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        var vm = await spService.ClientesObtenerAsync(resolvedNegocioId.Value, id);
        if (vm is null) return NotFound();
        CompletarNombresDesdeCampoGeneral(vm);
        vm.NegocioNombre = baseVm.NegocioNombre;
        vm.RolActual = baseVm.RolActual;
        InicializarTelefonoParaVista(vm);
        await CargarCombosClienteAsync(vm);
        return View(vm);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(ClienteFormViewModel model)
    {
        var baseVm = await ObtenerBaseAsync(model.NegocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEditar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });
        NormalizarYValidarIdentidad(model);
        ComponerTelefono(model);
        await NormalizarYValidarUbigeoAsync(model);
        if (!ModelState.IsValid)
        {
            await CargarCombosClienteAsync(model);
            return View(model);
        }

        try
        {
            var ok = await spService.ClientesActualizarAsync(model, User.Identity?.Name ?? "sistema");
            if (!ok)
            {
                ModelState.AddModelError(string.Empty, "No se pudo guardar el cliente. Verifica el negocio seleccionado.");
                await CargarCombosClienteAsync(model);
                return View(model);
            }
            return RedirectToAction(nameof(Index), new { negocioId = model.NegocioId });
        }
        catch (SqlException ex) when (EsErrorClienteDuplicado(ex.Message))
        {
            ModelState.AddModelError(string.Empty, "Cliente ya se encuentra registrado.");
            await CargarCombosClienteAsync(model);
            return View(model);
        }
        catch (SqlException ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            await CargarCombosClienteAsync(model);
            return View(model);
        }
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoProvincias(string? codigoDepartamento)
    {
        var codigoDep = (codigoDepartamento ?? string.Empty).Trim();
        if (codigoDep.Length != 2)
            return Json(Array.Empty<object>());

        var data = await spService.UbigeoProvinciasListarAsync(codigoDep);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
    }

    [HttpGet]
    public async Task<IActionResult> UbigeoDistritos(string? codigoProvincia)
    {
        var codigoProv = (codigoProvincia ?? string.Empty).Trim();
        if (codigoProv.Length != 4)
            return Json(Array.Empty<object>());

        var data = await spService.UbigeoDistritosListarAsync(codigoProv);
        return Json(data.Select(x => new { value = x.Value, text = x.Text }));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Delete(int negocioId, int id)
    {
        var baseVm = await ObtenerBaseAsync(negocioId, "CLIENTES");
        if (baseVm is null || !baseVm.PuedeEliminar) return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "No autorizado." });

        try
        {
            var ok = await spService.ClientesEliminarAsync(negocioId, id, User.Identity?.Name ?? "sistema");
            if (!ok) return NotFound();
            TempData["ClientesOk"] = "Cliente inactivado correctamente.";
        }
        catch (SqlException ex)
        {
            TempData["ClientesError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { negocioId, estado = "activos" });
    }

    private static void InicializarTelefonoParaVista(ClienteFormViewModel model)
    {
        TelefonoInternacionalHelper.Descomponer(model.Telefono, out var codigoPais, out var numeroLocal);
        model.TelefonoCodigoPais = codigoPais;
        model.TelefonoNumeroLocal = numeroLocal;
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
    }

    private static void ComponerTelefono(ClienteFormViewModel model)
    {
        model.Telefono = TelefonoInternacionalHelper.Componer(model.TelefonoCodigoPais, model.TelefonoNumeroLocal);
    }

    private static bool EsDocumentoRuc(string? tipoDocumento)
    {
        var codigo = (tipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        return codigo == CodigoDocumentoRucSunat || codigo == "RUC";
    }

    private static bool EsDocumentoNoDomiciliadoSinRuc(string? tipoDocumento)
    {
        var codigo = (tipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        return codigo == CodigoDocumentoNoDomiciliadoSinRucSunat;
    }

    private static void CompletarNombresDesdeCampoGeneral(ClienteFormViewModel model)
    {
        if (EsDocumentoRuc(model.TipoDocumento))
        {
            model.Nombres = null;
            model.Apellidos = null;
            return;
        }

        if (!string.IsNullOrWhiteSpace(model.Nombres) || !string.IsNullOrWhiteSpace(model.Apellidos))
            return;

        var partes = (model.NombresORazonSocial ?? string.Empty)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (partes.Length <= 1)
        {
            model.Nombres = model.NombresORazonSocial;
            return;
        }

        model.Nombres = string.Join(' ', partes.Take(partes.Length - 1));
        model.Apellidos = partes[^1];
    }

    private void NormalizarYValidarIdentidad(ClienteFormViewModel model)
    {
        model.TipoDocumento = (model.TipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        model.NombresORazonSocial = string.IsNullOrWhiteSpace(model.NombresORazonSocial) ? string.Empty : model.NombresORazonSocial.Trim();
        model.Nombres = string.IsNullOrWhiteSpace(model.Nombres) ? null : model.Nombres.Trim();
        model.Apellidos = string.IsNullOrWhiteSpace(model.Apellidos) ? null : model.Apellidos.Trim();
        model.NumeroDocumento = string.IsNullOrWhiteSpace(model.NumeroDocumento) ? string.Empty : model.NumeroDocumento.Trim();

        if (EsDocumentoNoDomiciliadoSinRuc(model.TipoDocumento))
        {
            model.NumeroDocumento = string.Empty;
        }
        else
        {
            if (string.IsNullOrWhiteSpace(model.NumeroDocumento))
                ModelState.AddModelError(nameof(model.NumeroDocumento), "Ingresa el numero de documento.");
            else
            {
                if (model.NumeroDocumento.Length > 11)
                    ModelState.AddModelError(nameof(model.NumeroDocumento), "El numero de documento permite como maximo 11 digitos.");
                if (!model.NumeroDocumento.All(char.IsDigit))
                    ModelState.AddModelError(nameof(model.NumeroDocumento), "El numero de documento solo permite digitos.");
            }
        }

        if (EsDocumentoRuc(model.TipoDocumento))
        {
            if (string.IsNullOrWhiteSpace(model.NombresORazonSocial))
                ModelState.AddModelError(nameof(model.NombresORazonSocial), "Ingresa la razon social.");

            model.Nombres = null;
            model.Apellidos = null;
            return;
        }

        if (string.IsNullOrWhiteSpace(model.Nombres))
            ModelState.AddModelError(nameof(model.Nombres), "Ingresa los nombres.");
        if (string.IsNullOrWhiteSpace(model.Apellidos))
            ModelState.AddModelError(nameof(model.Apellidos), "Ingresa los apellidos.");

        model.NombresORazonSocial = $"{model.Nombres} {model.Apellidos}".Trim();
    }

    private static bool EsErrorClienteDuplicado(string? mensaje)
    {
        return !string.IsNullOrWhiteSpace(mensaje) &&
               mensaje.Contains("Cliente ya se encuentra registrado", StringComparison.OrdinalIgnoreCase);
    }

    private async Task CargarCombosClienteAsync(ClienteFormViewModel model)
    {
        model.CodigosPais = TelefonoInternacionalHelper.ObtenerCodigosPais(model.TelefonoCodigoPais);
        model.TiposDocumento = await spService.CombosTiposDocumentoIdentidadSunatAsync();
        model.DepartamentosUbigeo = await spService.UbigeoDepartamentosListarAsync();

        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo) && Regex.IsMatch(model.CodigoUbigeo, @"^\d{6}$"))
        {
            model.CodigoDepartamento = model.CodigoUbigeo[..2];
            model.CodigoProvincia = model.CodigoUbigeo[..4];
        }

        model.ProvinciasUbigeo = !string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length == 2
            ? await spService.UbigeoProvinciasListarAsync(model.CodigoDepartamento)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();

        model.DistritosUbigeo = !string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length == 4
            ? await spService.UbigeoDistritosListarAsync(model.CodigoProvincia)
            : new List<Microsoft.AspNetCore.Mvc.Rendering.SelectListItem>();
    }

    private async Task NormalizarYValidarUbigeoAsync(ClienteFormViewModel model)
    {
        model.DireccionFiscal = string.IsNullOrWhiteSpace(model.DireccionFiscal) ? null : model.DireccionFiscal.Trim();
        model.CodigoDepartamento = string.IsNullOrWhiteSpace(model.CodigoDepartamento) ? null : model.CodigoDepartamento.Trim();
        model.CodigoProvincia = string.IsNullOrWhiteSpace(model.CodigoProvincia) ? null : model.CodigoProvincia.Trim();
        model.CodigoUbigeo = string.IsNullOrWhiteSpace(model.CodigoUbigeo) ? null : model.CodigoUbigeo.Trim();

        if (string.IsNullOrWhiteSpace(model.DireccionFiscal))
        {
            model.CodigoDepartamento = null;
            model.CodigoProvincia = null;
            model.CodigoUbigeo = null;
            await CargarCombosClienteAsync(model);
            return;
        }

        if (string.IsNullOrWhiteSpace(model.CodigoDepartamento))
            ModelState.AddModelError(nameof(model.CodigoDepartamento), "Selecciona un departamento.");
        if (string.IsNullOrWhiteSpace(model.CodigoProvincia))
            ModelState.AddModelError(nameof(model.CodigoProvincia), "Selecciona una provincia.");
        if (string.IsNullOrWhiteSpace(model.CodigoUbigeo))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Selecciona un distrito.");

        if (!string.IsNullOrWhiteSpace(model.CodigoDepartamento) && model.CodigoDepartamento.Length != 2)
            ModelState.AddModelError(nameof(model.CodigoDepartamento), "Codigo de departamento invalido.");
        if (!string.IsNullOrWhiteSpace(model.CodigoProvincia) && model.CodigoProvincia.Length != 4)
            ModelState.AddModelError(nameof(model.CodigoProvincia), "Codigo de provincia invalido.");
        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo) && !Regex.IsMatch(model.CodigoUbigeo, @"^\d{6}$"))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "Codigo de distrito invalido.");

        if (!string.IsNullOrWhiteSpace(model.CodigoDepartamento) &&
            !string.IsNullOrWhiteSpace(model.CodigoProvincia) &&
            !model.CodigoProvincia.StartsWith(model.CodigoDepartamento, StringComparison.Ordinal))
            ModelState.AddModelError(nameof(model.CodigoProvincia), "La provincia no corresponde al departamento seleccionado.");

        if (!string.IsNullOrWhiteSpace(model.CodigoProvincia) &&
            !string.IsNullOrWhiteSpace(model.CodigoUbigeo) &&
            !model.CodigoUbigeo.StartsWith(model.CodigoProvincia, StringComparison.Ordinal))
            ModelState.AddModelError(nameof(model.CodigoUbigeo), "El distrito no corresponde a la provincia seleccionada.");

        if (!string.IsNullOrWhiteSpace(model.CodigoUbigeo))
        {
            var ubigeo = await spService.UbigeoObtenerPorCodigoAsync(model.CodigoUbigeo);
            if (ubigeo is null)
            {
                ModelState.AddModelError(nameof(model.CodigoUbigeo), "El distrito seleccionado no existe.");
            }
            else
            {
                model.CodigoDepartamento = ubigeo.CodigoDepartamento;
                model.CodigoProvincia = ubigeo.CodigoProvincia;
            }
        }

        await CargarCombosClienteAsync(model);
    }
}
