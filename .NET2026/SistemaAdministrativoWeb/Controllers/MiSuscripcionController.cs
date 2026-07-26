using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SistemaAdministrativoWeb.Infrastructure.Empresas;
using SistemaAdministrativoWeb.Infrastructure.Security;
using SistemaAdministrativoWeb.Infrastructure.Suscripciones;
using SistemaAdministrativoWeb.ViewModels.Suscripciones;

namespace SistemaAdministrativoWeb.Controllers;

[Authorize]
[ModulePermission("MISUSCRIPCION")]
[AllowRestrictedSubscription]
public class MiSuscripcionController(
    ICurrentCompanyAccessor currentCompanyAccessor,
    ISubscriptionAccessService subscriptionAccessService) : Controller
{
    [HttpGet]
    public async Task<IActionResult> Index(CancellationToken cancellationToken)
    {
        if (User.IsInRole("SuperAdmin"))
        {
            return RedirectToAction("Index", "Plataforma");
        }

        ViewData["AdminShell"] = true;

        var evaluacion = await subscriptionAccessService.EvaluateAsync(User, cancellationToken);
        var contexto = evaluacion.LoginContext;

        if (contexto is null || !contexto.TieneAcceso || !contexto.IdCuentaAdministradora.HasValue)
        {
            return RedirectToAction("Index", "EmpresaContexto");
        }

        var model = new MiSuscripcionIndexViewModel
        {
            IdEmpresa = currentCompanyAccessor.EmpresaId ?? contexto.IdEmpresaPredeterminada ?? 0,
            NombreEmpresa = currentCompanyAccessor.EmpresaNombre ?? contexto.RazonSocialEmpresaPredeterminada ?? "Empresa",
            IdCuentaAdministradora = contexto.IdCuentaAdministradora.Value,
            CodigoCuenta = contexto.CodigoCuenta ?? string.Empty,
            NombreCuenta = contexto.NombreCuenta ?? "Cuenta administradora",
            CorreoPrincipal = contexto.CorreoPrincipal ?? string.Empty,
            TelefonoPrincipal = contexto.TelefonoPrincipal,
            TipoPlan = evaluacion.PlanDisplay,
            EstadoSuscripcion = evaluacion.StatusDisplay,
            EsPrueba = contexto.EsPrueba == true,
            FechaInicioPrueba = contexto.FechaInicioPrueba,
            FechaFinPrueba = contexto.FechaFinPrueba,
            FechaInicioPlan = contexto.FechaInicioPlan,
            FechaFinPlan = contexto.FechaFinPlan,
            FechaFinGracia = evaluacion.GraceEndDate,
            EmpresasPermitidas = contexto.EmpresasPermitidas,
            UsuariosPermitidos = contexto.UsuariosPermitidos,
            EstadoCuenta = contexto.EstadoCuenta == true,
            EstadoEmpresa = true,
            Observacion = contexto.ObservacionSuscripcion,
            FechaVencimiento = evaluacion.ExpirationDate,
            DiasParaVencer = evaluacion.DaysUntilExpiration,
            ResumenVigencia = evaluacion.Message,
            EstadoEfectivo = evaluacion.EffectiveStatus,
            AccesoRestringido = evaluacion.IsRestricted
        };

        return View(model);
    }
}
