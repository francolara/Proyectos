using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Plataforma;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize(Roles = "SuperAdmin")]
public class PlataformaController(ICuentaAdministradoraRepository cuentaAdministradoraRepository) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        ViewData["PlatformShell"] = true;

        var cuentas = await cuentaAdministradoraRepository.ListarCuentasSuscripcionAsync(cancellationToken);
        var model = new PlataformaIndexViewModel
        {
            TotalCuentas = cuentas.Count,
            CuentasActivas = cuentas.Count(x => x.Activo && x.EstadoCuenta && string.Equals(x.EstadoSuscripcion, "ACTIVO", StringComparison.OrdinalIgnoreCase)),
            CuentasEnPrueba = cuentas.Count(x => x.Activo && x.EstadoCuenta && string.Equals(x.EstadoSuscripcion, "TRIAL", StringComparison.OrdinalIgnoreCase)),
            CuentasSuspendidasOBaja = cuentas.Count(x => !x.Activo || !x.EstadoCuenta || string.Equals(x.EstadoSuscripcion, "SUSPENDIDO", StringComparison.OrdinalIgnoreCase) || string.Equals(x.EstadoSuscripcion, "BAJA", StringComparison.OrdinalIgnoreCase)),
            Cuentas = cuentas.Select(x => new CuentaSuscripcionViewModel
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
                EsPrueba = x.EsPrueba,
                FechaInicioPrueba = x.FechaInicioPrueba,
                FechaFinPrueba = x.FechaFinPrueba,
                FechaInicioPlan = x.FechaInicioPlan,
                FechaFinPlan = x.FechaFinPlan,
                EmpresasPermitidas = x.EmpresasPermitidas,
                UsuariosPermitidos = x.UsuariosPermitidos,
                Activo = x.Activo,
                EstadoCuenta = x.EstadoCuenta,
                Observacion = x.Observacion
            }).ToList()
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
}
