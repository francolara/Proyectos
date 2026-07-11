using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Contabilidad;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Configuracion;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("CONFIGURACION")]
public class ConfiguracionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository,
    IMigoPadronApiClient migoPadronApiClient) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["Title"] = "Configuracion";
        ViewData["AdminShell"] = true;

        return View(await ConstruirModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken));
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Guardar(ConfiguracionCuentaAdministradoraViewModel model, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ValidarModelo(model);
        if (!ModelState.IsValid)
        {
            ViewData["Title"] = "Configuracion";
            ViewData["AdminShell"] = true;
            var recargado = await ConstruirModelAsync(cuenta.Value.idCuentaAdministradora, cancellationToken);
            model.CodigoCuenta = recargado.CodigoCuenta;
            model.NombreCuenta = recargado.NombreCuenta;
            model.CorreoPrincipal = recargado.CorreoPrincipal;
            model.TelefonoPrincipal = recargado.TelefonoPrincipal;
            model.EmpresasDisponibles = recargado.EmpresasDisponibles;
            return View("Index", model);
        }

        await cuentaAdministradoraRepository.GuardarConfiguracionCuentaAdministradoraAsync(new GuardarConfiguracionCuentaAdministradoraRequest
        {
            IdCuentaAdministradora = cuenta.Value.idCuentaAdministradora,
            NombreResponsablePrincipal = LimpiarTexto(model.NombreResponsablePrincipal),
            CorreoAdministrativo = LimpiarTexto(model.CorreoAdministrativo),
            TelefonoAdministrativo = LimpiarTelefono(model.TelefonoAdministrativo),
            IdEmpresaPredeterminada = model.IdEmpresaPredeterminada,
            ObservacionAdministrativa = LimpiarTexto(model.ObservacionAdministrativa),
            TipoComprobantePreferido = model.TipoComprobantePreferido.Trim().ToUpperInvariant(),
            TipoDocumentoFacturacion = model.TipoDocumentoFacturacion.Trim().ToUpperInvariant(),
            NumeroDocumento = LimpiarDocumento(model.NumeroDocumento),
            NombreFacturacion = LimpiarTexto(model.NombreFacturacion),
            RazonSocialFacturacion = LimpiarTexto(model.RazonSocialFacturacion),
            CorreoFacturacion = LimpiarTexto(model.CorreoFacturacion),
            TelefonoFacturacion = LimpiarTelefono(model.TelefonoFacturacion),
            DireccionFiscal = LimpiarTexto(model.DireccionFiscal),
            Ubigeo = LimpiarTexto(model.Ubigeo),
            Distrito = LimpiarTexto(model.Distrito),
            Provincia = LimpiarTexto(model.Provincia),
            Departamento = LimpiarTexto(model.Departamento),
            ObservacionFacturacion = LimpiarTexto(model.ObservacionFacturacion),
            UsuarioRegistro = User.Identity?.Name
        }, cancellationToken);

        TempData["SuccessMessage"] = "La configuracion de la cuenta administradora fue actualizada.";
        return RedirectToAction(nameof(Index));
    }

    [HttpGet]
    public async Task<IActionResult> ConsultarPadron(string tipoDocumento, string numeroDocumento, CancellationToken cancellationToken)
    {
        var cuenta = await ResolverCuentaAsync(cancellationToken);
        if (cuenta is null)
        {
            return Unauthorized();
        }

        var tipo = (tipoDocumento ?? string.Empty).Trim().ToUpperInvariant();
        var numero = LimpiarDocumento(numeroDocumento);
        if (string.IsNullOrWhiteSpace(numero))
        {
            return Json(new FacturacionPadronLookupResultViewModel
            {
                Encontrado = false,
                Mensaje = "Ingrese un numero de documento valido."
            });
        }

        if (tipo == "RUC")
        {
            var result = await migoPadronApiClient.ConsultarRucAsync(numero, cancellationToken);
            if (result is null)
            {
                return Json(new FacturacionPadronLookupResultViewModel
                {
                    Encontrado = false,
                    Mensaje = "No se encontro informacion para el RUC ingresado."
                });
            }

            return Json(new FacturacionPadronLookupResultViewModel
            {
                Encontrado = true,
                NumeroDocumento = result.Ruc,
                RazonSocialFacturacion = result.NombreORazonSocial,
                DireccionFiscal = result.DireccionSimple ?? result.Direccion,
                Ubigeo = result.Ubigeo,
                Distrito = result.Distrito,
                Provincia = result.Provincia,
                Departamento = result.Departamento,
                Mensaje = "Datos obtenidos desde Migo."
            });
        }

        if (tipo == "DNI")
        {
            var result = await migoPadronApiClient.ConsultarDniAsync(numero, cancellationToken);
            if (result is null)
            {
                return Json(new FacturacionPadronLookupResultViewModel
                {
                    Encontrado = false,
                    Mensaje = "No se encontro informacion para el DNI ingresado."
                });
            }

            return Json(new FacturacionPadronLookupResultViewModel
            {
                Encontrado = true,
                NumeroDocumento = result.Dni,
                NombreFacturacion = result.NombreCompleto,
                Mensaje = "Datos obtenidos desde Migo."
            });
        }

        return Json(new FacturacionPadronLookupResultViewModel
        {
            Encontrado = false,
            Mensaje = "La consulta automatica solo esta disponible para DNI y RUC."
        });
    }

    private async Task<ConfiguracionCuentaAdministradoraViewModel> ConstruirModelAsync(int idCuentaAdministradora, CancellationToken cancellationToken)
    {
        var configuracion = await cuentaAdministradoraRepository.ObtenerConfiguracionCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);
        var empresas = await cuentaAdministradoraRepository.ListarEmpresasCuentaAdministradoraAsync(idCuentaAdministradora, cancellationToken);
        var correoPrincipal = configuracion?.CorreoPrincipal;
        var telefonoPrincipal = configuracion?.TelefonoPrincipal;

        return new ConfiguracionCuentaAdministradoraViewModel
        {
            IdCuentaAdministradora = idCuentaAdministradora,
            CodigoCuenta = configuracion?.CodigoCuenta ?? string.Empty,
            NombreCuenta = configuracion?.NombreCuenta ?? string.Empty,
            CorreoPrincipal = correoPrincipal,
            TelefonoPrincipal = telefonoPrincipal,
            NombreResponsablePrincipal = configuracion?.NombreResponsablePrincipal,
            CorreoAdministrativo = string.IsNullOrWhiteSpace(configuracion?.CorreoAdministrativo)
                ? correoPrincipal
                : configuracion!.CorreoAdministrativo,
            TelefonoAdministrativo = string.IsNullOrWhiteSpace(configuracion?.TelefonoAdministrativo)
                ? telefonoPrincipal
                : configuracion!.TelefonoAdministrativo,
            IdEmpresaPredeterminada = configuracion?.IdEmpresaPredeterminada,
            ObservacionAdministrativa = configuracion?.ObservacionAdministrativa,
            TipoComprobantePreferido = configuracion?.TipoComprobantePreferido ?? "BOLETA",
            TipoDocumentoFacturacion = configuracion?.TipoDocumentoFacturacion ?? "DNI",
            NumeroDocumento = configuracion?.NumeroDocumento,
            NombreFacturacion = configuracion?.NombreFacturacion,
            RazonSocialFacturacion = configuracion?.RazonSocialFacturacion,
            CorreoFacturacion = configuracion?.CorreoFacturacion,
            TelefonoFacturacion = configuracion?.TelefonoFacturacion,
            DireccionFiscal = configuracion?.DireccionFiscal,
            Ubigeo = configuracion?.Ubigeo,
            Distrito = configuracion?.Distrito,
            Provincia = configuracion?.Provincia,
            Departamento = configuracion?.Departamento,
            ObservacionFacturacion = configuracion?.ObservacionFacturacion,
            EmpresasDisponibles = empresas
                .Select(x => new ConfiguracionEmpresaItemViewModel
                {
                    IdEmpresa = x.IdEmpresa,
                    CodigoEmpresa = x.CodigoEmpresa,
                    RazonSocial = x.RazonSocial
                })
                .ToList()
        };
    }

    private async Task<(int idCuentaAdministradora, string? nombreCuenta)?> ResolverCuentaAsync(CancellationToken cancellationToken)
    {
        if (currentCompanyAccessor.TieneEmpresaActiva && currentCompanyAccessor.EmpresaId.HasValue)
        {
            var contextoEmpresa = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(
                currentCompanyAccessor.EmpresaId.Value,
                cancellationToken);

            if (contextoEmpresa is not null)
            {
                return (contextoEmpresa.IdCuentaAdministradora, contextoEmpresa.NombreCuenta);
            }
        }

        var aspNetUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);
        if (string.IsNullOrWhiteSpace(aspNetUserId))
        {
            return null;
        }

        var contextoLogin = await cuentaAdministradoraRepository.ObtenerContextoLoginUsuarioAsync(aspNetUserId, cancellationToken);
        if (contextoLogin is null || !contextoLogin.TieneAcceso || !contextoLogin.IdCuentaAdministradora.HasValue)
        {
            return null;
        }

        return (contextoLogin.IdCuentaAdministradora.Value, contextoLogin.NombreCuenta);
    }

    private void ValidarModelo(ConfiguracionCuentaAdministradoraViewModel model)
    {
        model.TipoComprobantePreferido = (model.TipoComprobantePreferido ?? "BOLETA").Trim().ToUpperInvariant();
        model.TipoDocumentoFacturacion = (model.TipoDocumentoFacturacion ?? "DNI").Trim().ToUpperInvariant();

        var numeroDocumento = LimpiarDocumento(model.NumeroDocumento);
        if (model.TipoComprobantePreferido == "FACTURA" && model.TipoDocumentoFacturacion != "RUC")
        {
            ModelState.AddModelError(nameof(model.TipoDocumentoFacturacion), "Para factura debe registrarse un RUC.");
        }

        if (model.TipoDocumentoFacturacion == "RUC" && !string.IsNullOrWhiteSpace(numeroDocumento) && numeroDocumento.Length != 11)
        {
            ModelState.AddModelError(nameof(model.NumeroDocumento), "El RUC debe tener 11 digitos.");
        }

        if (model.TipoDocumentoFacturacion == "DNI" && !string.IsNullOrWhiteSpace(numeroDocumento) && numeroDocumento.Length != 8)
        {
            ModelState.AddModelError(nameof(model.NumeroDocumento), "El DNI debe tener 8 digitos.");
        }

        if (model.TipoComprobantePreferido == "BOLETA" && string.IsNullOrWhiteSpace(model.NombreFacturacion))
        {
            ModelState.AddModelError(nameof(model.NombreFacturacion), "Ingrese el nombre a usar en la boleta.");
        }

        if (model.TipoComprobantePreferido == "FACTURA" && string.IsNullOrWhiteSpace(model.RazonSocialFacturacion))
        {
            ModelState.AddModelError(nameof(model.RazonSocialFacturacion), "Ingrese la razon social a usar en la factura.");
        }
    }

    private static string? LimpiarTelefono(string? telefono)
    {
        if (string.IsNullOrWhiteSpace(telefono))
        {
            return null;
        }

        return new string(telefono.Where(x => char.IsDigit(x) || x == '+').ToArray());
    }

    private static string? LimpiarDocumento(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return new string(value.Where(char.IsLetterOrDigit).ToArray()).ToUpperInvariant();
    }

    private static string? LimpiarTexto(string? value)
        => string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
