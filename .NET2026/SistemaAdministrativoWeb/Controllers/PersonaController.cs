using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.ViewModels.Contabilidad;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("PERSONAS")]
public class PersonaController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    IPersonaRepository personaRepository,
    IMigoPadronApiClient migoPadronApiClient) : Controller
{
    private const int TamanoPagina = 20;

    [HttpGet]
    public async Task<IActionResult> Index(string? textoBusqueda = null, string? tipoPersona = null, bool soloClientes = false, bool soloProveedores = false, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var personas = await personaRepository.ListarPaginadoPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            textoBusqueda,
            tipoPersona,
            soloClientes,
            soloProveedores,
            pagina,
            TamanoPagina,
            cancellationToken);

        var model = await ConstruirListadoAsync(personas.Items, cancellationToken);
        model.TextoBusqueda = textoBusqueda?.Trim() ?? string.Empty;
        model.TipoPersonaFiltro = tipoPersona?.Trim().ToUpperInvariant() ?? string.Empty;
        model.SoloClientes = soloClientes;
        model.SoloProveedores = soloProveedores;
        model.TotalPersonas = personas.TotalRecords;
        model.Paginacion = new PaginacionViewModel
        {
            PaginaActual = pagina,
            TamanoPagina = TamanoPagina,
            TotalRegistros = personas.TotalRecords
        };

        return View(model);
    }

    [HttpGet]
    public async Task<IActionResult> Registrar(CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(null, cancellationToken);
    }

    [HttpGet]
    public async Task<IActionResult> Editar(int idPersona, CancellationToken cancellationToken = default)
    {
        return await CargarFormularioAsync(idPersona, cancellationToken);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModulePermission("PERSONAS", ModulePermissionOperation.Delete)]
    public async Task<IActionResult> Eliminar(int idPersona, string? textoBusqueda = null, string? tipoPersona = null, bool soloClientes = false, bool soloProveedores = false, int pagina = 1, CancellationToken cancellationToken = default)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        try
        {
            await personaRepository.EliminarAsync(currentCompanyAccessor.EmpresaId.Value, idPersona, cancellationToken);
            TempData["PersonaOk"] = "Persona eliminada correctamente.";
        }
        catch (Exception ex)
        {
            TempData["PersonaError"] = ex.Message;
        }

        return RedirectToAction(nameof(Index), new { textoBusqueda, tipoPersona, soloClientes, soloProveedores, pagina });
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    [ModuleSavePermission("PERSONAS", nameof(PersonaFormViewModel.IdPersona))]
    public async Task<IActionResult> Guardar(PersonaFormViewModel formulario, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;
        ValidarFormulario(formulario);

        if (!ModelState.IsValid)
        {
            var modelConError = await ConstruirFormularioAsync(formulario, cancellationToken);
            return View("Formulario", modelConError);
        }

        try
        {
            await personaRepository.GuardarAsync(new GuardarPersonaRequest
            {
                IdPersona = formulario.IdPersona,
                IdEmpresa = currentCompanyAccessor.EmpresaId.Value,
                TipoPersona = formulario.TipoPersona.Trim().ToUpperInvariant(),
                TipoDocumento = formulario.TipoDocumento.Trim().ToUpperInvariant(),
                NumeroDocumento = formulario.NumeroDocumento.Trim(),
                ApellidoPaterno = formulario.ApellidoPaterno,
                ApellidoMaterno = formulario.ApellidoMaterno,
                Nombres = formulario.Nombres,
                RazonSocial = formulario.RazonSocial,
                CorreoElectronico = formulario.CorreoElectronico,
                Telefono = formulario.Telefono,
                Direccion = formulario.Direccion,
                CodigoUbigeo = formulario.CodigoUbigeo,
                EsCliente = formulario.EsCliente,
                EsProveedor = formulario.EsProveedor,
                Estado = formulario.Estado,
                UsuarioRegistro = User.Identity?.Name
            }, cancellationToken);

            TempData["PersonaOk"] = formulario.IdPersona.HasValue
                ? "Persona actualizada correctamente."
                : "Persona registrada correctamente.";

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            ModelState.AddModelError(string.Empty, ex.Message);
            var modelConError = await ConstruirFormularioAsync(formulario, cancellationToken);
            return View("Formulario", modelConError);
        }
    }

    [HttpGet]
    public async Task<IActionResult> Provincias(string codigoDepartamento, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(codigoDepartamento))
        {
            return Json(Array.Empty<object>());
        }

        var provincias = await personaRepository.ListarProvinciasAsync(codigoDepartamento.Trim(), cancellationToken);
        return Json(provincias.Select(x => new { value = x.CodigoProvincia, text = x.Nombre }));
    }

    [HttpGet]
    public async Task<IActionResult> Distritos(string codigoProvincia, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(codigoProvincia))
        {
            return Json(Array.Empty<object>());
        }

        var distritos = await personaRepository.ListarDistritosAsync(codigoProvincia.Trim(), cancellationToken);
        return Json(distritos.Select(x => new { value = x.CodigoUbigeo, text = x.Nombre }));
    }

    [HttpGet]
    public async Task<IActionResult> ConsultarDocumentoMigo(string? tipoDocumento = null, string? numeroDocumento = null, CancellationToken cancellationToken = default)
    {
        var tipoDocumentoTrabajo = (tipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        var numeroDocumentoTrabajo = new string((numeroDocumento ?? string.Empty).Where(char.IsDigit).ToArray());

        if (string.IsNullOrWhiteSpace(tipoDocumentoTrabajo) || string.IsNullOrWhiteSpace(numeroDocumentoTrabajo))
        {
            return Json(new { ok = false, mensaje = "Seleccione el tipo de documento e ingrese el numero antes de consultar." });
        }

        try
        {
            if (tipoDocumentoTrabajo == "6")
            {
                var ruc = await migoPadronApiClient.ConsultarRucAsync(numeroDocumentoTrabajo, cancellationToken);
                if (ruc is null)
                {
                    return Json(new { ok = false, mensaje = "La API no devolvio informacion para el RUC consultado." });
                }

                return Json(new
                {
                    ok = true,
                    tipoPersona = "J",
                    tipoDocumento = "6",
                    numeroDocumento = ruc.Ruc,
                    razonSocial = ruc.NombreORazonSocial,
                    direccion = ruc.DireccionSimple ?? ruc.Direccion ?? string.Empty,
                    codigoUbigeo = ruc.Ubigeo ?? string.Empty,
                    codigoDepartamento = ObtenerCodigoDepartamento(ruc.Ubigeo),
                    codigoProvincia = ObtenerCodigoProvincia(ruc.Ubigeo),
                    distrito = ruc.Distrito ?? string.Empty,
                    provincia = ruc.Provincia ?? string.Empty,
                    departamento = ruc.Departamento ?? string.Empty,
                    estadoContribuyente = ruc.EstadoContribuyente ?? string.Empty,
                    condicionDomicilio = ruc.CondicionDomicilio ?? string.Empty
                });
            }

            if (tipoDocumentoTrabajo == "1")
            {
                var dni = await migoPadronApiClient.ConsultarDniAsync(numeroDocumentoTrabajo, cancellationToken);
                if (dni is null)
                {
                    return Json(new { ok = false, mensaje = "La API no devolvio informacion para el DNI consultado." });
                }

                var nombres = SepararNombreDni(dni.NombreCompleto);
                return Json(new
                {
                    ok = true,
                    tipoPersona = "N",
                    tipoDocumento = "1",
                    numeroDocumento = dni.Dni,
                    apellidoPaterno = nombres.apellidoPaterno,
                    apellidoMaterno = nombres.apellidoMaterno,
                    nombres = nombres.nombres,
                    nombreCompleto = dni.NombreCompleto
                });
            }

            return Json(new { ok = false, mensaje = "Solo se puede consultar RUC o DNI desde la API de Migo." });
        }
        catch (Exception ex)
        {
            return Json(new { ok = false, mensaje = ex.Message });
        }
    }

    private async Task<IActionResult> CargarFormularioAsync(int? idPersona, CancellationToken cancellationToken)
    {
        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        PersonaFormViewModel formulario;
        if (idPersona.HasValue)
        {
            var persona = await personaRepository.ObtenerPorIdAsync(currentCompanyAccessor.EmpresaId.Value, idPersona.Value, cancellationToken);
            if (persona is null)
            {
                return RedirectToAction(nameof(Index));
            }

            formulario = new PersonaFormViewModel
            {
                IdPersona = persona.IdPersona,
                TipoPersona = persona.TipoPersona,
                TipoDocumento = persona.TipoDocumento,
                NumeroDocumento = persona.NumeroDocumento,
                ApellidoPaterno = persona.ApellidoPaterno,
                ApellidoMaterno = persona.ApellidoMaterno,
                Nombres = persona.Nombres,
                RazonSocial = persona.RazonSocial,
                CorreoElectronico = persona.CorreoElectronico,
                Telefono = persona.Telefono,
                Direccion = persona.Direccion,
                CodigoDepartamento = persona.CodigoDepartamento,
                CodigoProvincia = persona.CodigoProvincia,
                CodigoUbigeo = persona.CodigoUbigeo,
                EsCliente = persona.EsCliente,
                EsProveedor = persona.EsProveedor,
                Estado = persona.Estado
            };
        }
        else
        {
            formulario = new PersonaFormViewModel();
        }

        var model = await ConstruirFormularioAsync(formulario, cancellationToken);
        return View("Formulario", model);
    }

    private async Task<PersonaIndexViewModel> ConstruirListadoAsync(IReadOnlyCollection<PersonaDto> personas, CancellationToken cancellationToken)
    {
        var model = await ConstruirBaseAsync(cancellationToken);
        model.Personas = personas
            .Select(x => new PersonaItemViewModel
            {
                IdPersona = x.IdPersona,
                TipoPersona = x.TipoPersona,
                TipoDocumento = x.TipoDocumento,
                NombreTipoDocumento = x.NombreTipoDocumento,
                NumeroDocumento = x.NumeroDocumento,
                NombreCompleto = x.NombreCompleto,
                CorreoElectronico = x.CorreoElectronico,
                Telefono = x.Telefono,
                Direccion = x.Direccion,
                Departamento = x.Departamento,
                Provincia = x.Provincia,
                Distrito = x.Distrito,
                EsCliente = x.EsCliente,
                EsProveedor = x.EsProveedor,
                Estado = x.Estado
            })
            .ToList();
        model.TotalClientes = model.Personas.Count(x => x.EsCliente);
        model.TotalProveedores = model.Personas.Count(x => x.EsProveedor);
        return model;
    }

    private async Task<PersonaIndexViewModel> ConstruirFormularioAsync(PersonaFormViewModel formulario, CancellationToken cancellationToken)
    {
        var model = await ConstruirBaseAsync(cancellationToken);
        model.Formulario = formulario;
        model.Provincias = await ListarProvinciasAsync(formulario.CodigoDepartamento, cancellationToken);
        model.Distritos = await ListarDistritosAsync(formulario.CodigoProvincia, cancellationToken);
        return model;
    }

    private async Task<PersonaIndexViewModel> ConstruirBaseAsync(CancellationToken cancellationToken)
    {
        var tiposDocumento = await personaRepository.ListarTiposDocumentoAsync(cancellationToken);
        var departamentos = await personaRepository.ListarDepartamentosAsync(cancellationToken);

        return new PersonaIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? 0,
            EmpresaNombre = currentCompanyAccessor.EmpresaNombre ?? "Empresa activa",
            TiposPersona =
            [
                new OpcionCatalogoViewModel { Valor = "N", Texto = "Natural" },
                new OpcionCatalogoViewModel { Valor = "J", Texto = "Juridica" }
            ],
            TiposDocumento = tiposDocumento
                .Select(x => new OpcionCatalogoViewModel
                {
                    Valor = x.CodigoSunat,
                    Texto = $"{x.CodigoSunat} - {x.Nombre}"
                })
                .ToList(),
            Departamentos = departamentos
                .Select(x => new OpcionCatalogoViewModel
                {
                    Valor = x.CodigoDepartamento,
                    Texto = x.Nombre
                })
                .ToList()
        };
    }

    private async Task<List<OpcionCatalogoViewModel>> ListarProvinciasAsync(string? codigoDepartamento, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(codigoDepartamento))
        {
            return [];
        }

        var provincias = await personaRepository.ListarProvinciasAsync(codigoDepartamento.Trim(), cancellationToken);
        return provincias
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoProvincia,
                Texto = x.Nombre
            })
            .ToList();
    }

    private async Task<List<OpcionCatalogoViewModel>> ListarDistritosAsync(string? codigoProvincia, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(codigoProvincia))
        {
            return [];
        }

        var distritos = await personaRepository.ListarDistritosAsync(codigoProvincia.Trim(), cancellationToken);
        return distritos
            .Select(x => new OpcionCatalogoViewModel
            {
                Valor = x.CodigoUbigeo,
                Texto = x.Nombre
            })
            .ToList();
    }

    private void ValidarFormulario(PersonaFormViewModel formulario)
    {
        if (string.Equals(formulario.TipoPersona, "N", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(formulario.Nombres))
            {
                ModelState.AddModelError("Formulario.Nombres", "Ingrese los nombres de la persona.");
            }
        }
        else if (string.Equals(formulario.TipoPersona, "J", StringComparison.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(formulario.RazonSocial))
            {
                ModelState.AddModelError("Formulario.RazonSocial", "Ingrese la razon social.");
            }
        }
    }

    private static string? ObtenerCodigoDepartamento(string? codigoUbigeo)
    {
        return !string.IsNullOrWhiteSpace(codigoUbigeo) && codigoUbigeo.Length >= 2
            ? codigoUbigeo[..2]
            : null;
    }

    private static string? ObtenerCodigoProvincia(string? codigoUbigeo)
    {
        return !string.IsNullOrWhiteSpace(codigoUbigeo) && codigoUbigeo.Length >= 4
            ? codigoUbigeo[..4]
            : null;
    }

    private static (string? apellidoPaterno, string? apellidoMaterno, string? nombres) SepararNombreDni(string nombreCompleto)
    {
        var partes = (nombreCompleto ?? string.Empty)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (partes.Length == 0)
        {
            return (null, null, null);
        }

        if (partes.Length == 1)
        {
            return (null, null, partes[0]);
        }

        if (partes.Length == 2)
        {
            return (partes[0], null, partes[1]);
        }

        return (partes[0], partes[1], string.Join(' ', partes.Skip(2)));
    }
}
