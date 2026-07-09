using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Suscripciones;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
public class MiSuscripcionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ICuentaAdministradoraRepository cuentaAdministradoraRepository) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        if (User.IsInRole("SuperAdmin") && !currentCompanyAccessor.TieneEmpresaActiva)
        {
            return RedirectToAction("Index", "Plataforma");
        }

        if (!currentCompanyAccessor.TieneEmpresaActiva || !currentCompanyAccessor.EmpresaId.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        ViewData["AdminShell"] = true;

        var contexto = await cuentaAdministradoraRepository.ObtenerContextoSuscripcionPorEmpresaAsync(
            currentCompanyAccessor.EmpresaId.Value,
            cancellationToken);

        if (contexto is null)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var fechaVencimiento = contexto.EsPrueba ? contexto.FechaFinPrueba : contexto.FechaFinPlan;
        int? diasParaVencer = null;
        if (fechaVencimiento.HasValue)
        {
            diasParaVencer = fechaVencimiento.Value.DayNumber - DateOnly.FromDateTime(DateTime.Today).DayNumber;
        }

        var model = new MiSuscripcionIndexViewModel
        {
            IdEmpresa = contexto.IdEmpresa,
            NombreEmpresa = currentCompanyAccessor.EmpresaNombre ?? contexto.RazonSocial,
            Ruc = contexto.Ruc,
            IdCuentaAdministradora = contexto.IdCuentaAdministradora,
            CodigoCuenta = contexto.CodigoCuenta,
            NombreCuenta = contexto.NombreCuenta,
            CorreoPrincipal = contexto.CorreoPrincipal,
            TelefonoPrincipal = contexto.TelefonoPrincipal,
            TipoPlan = string.IsNullOrWhiteSpace(contexto.TipoPlan) ? "TRIAL" : contexto.TipoPlan,
            EstadoSuscripcion = string.IsNullOrWhiteSpace(contexto.EstadoSuscripcion) ? "TRIAL" : contexto.EstadoSuscripcion,
            EsPrueba = contexto.EsPrueba,
            FechaInicioPrueba = contexto.FechaInicioPrueba,
            FechaFinPrueba = contexto.FechaFinPrueba,
            FechaInicioPlan = contexto.FechaInicioPlan,
            FechaFinPlan = contexto.FechaFinPlan,
            EmpresasPermitidas = contexto.EmpresasPermitidas,
            UsuariosPermitidos = contexto.UsuariosPermitidos,
            EstadoCuenta = contexto.EstadoCuenta,
            EstadoEmpresa = contexto.EstadoEmpresa,
            Observacion = contexto.Observacion,
            FechaVencimiento = fechaVencimiento,
            DiasParaVencer = diasParaVencer,
            ResumenVigencia = ResolverResumenVigencia(contexto, fechaVencimiento, diasParaVencer)
        };

        return View(model);
    }

    private static string ResolverResumenVigencia(ContextoSuscripcionEmpresaDto contexto, DateOnly? fechaVencimiento, int? diasParaVencer)
    {
        if (fechaVencimiento.HasValue)
        {
            if (diasParaVencer < 0)
            {
                return $"La suscripcion vencio el {fechaVencimiento:dd/MM/yyyy}.";
            }

            if (diasParaVencer <= 7)
            {
                return $"La suscripcion vence el {fechaVencimiento:dd/MM/yyyy}.";
            }

            return $"Cobertura vigente hasta {fechaVencimiento:dd/MM/yyyy}.";
        }

        if (!string.IsNullOrWhiteSpace(contexto.EstadoSuscripcion))
        {
            return $"Estado actual de la suscripcion: {contexto.EstadoSuscripcion}.";
        }

        return "Aun no se ha configurado una vigencia para esta cuenta administradora.";
    }
}
