using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

public class MiSuscripcionController(
    IModuloPermisoService moduloPermisoService,
    ISportCenterStoredProcedureService spService) : ModuloControllerBase(moduloPermisoService)
{
    public async Task<IActionResult> Index(int? negocioId)
    {
        var resolvedNegocioId = await ResolverNegocioIdAsync(negocioId, spService);
        if (!resolvedNegocioId.HasValue) return Forbid();

        var baseVm = await ObtenerBaseAsync(resolvedNegocioId.Value, "DASHBOARD");
        if (baseVm is null || !string.IsNullOrWhiteSpace(baseVm.Mensaje))
            return SinAcceso(baseVm ?? new ModuloBaseViewModel { Mensaje = "Acceso denegado." });

        var suscripcion = await spService.MiSuscripcionObtenerAsync(resolvedNegocioId.Value);
        var contactoEmail = (await spService.ParametrosGlobalesObtenerValorAsync("HOME_PORTAL_CONTACTO_EMAIL")) ?? string.Empty;
        var contactoTelefono = (await spService.ParametrosGlobalesObtenerValorAsync("HOME_PORTAL_CONTACTO_TELEFONO")) ?? string.Empty;

        DateTime? fechaVencimiento = null;
        if (suscripcion is not null)
            fechaVencimiento = suscripcion.EsPrueba ? suscripcion.FechaFinPrueba : suscripcion.FechaFinPlan;

        int? diasParaVencer = null;
        if (fechaVencimiento.HasValue)
            diasParaVencer = (int)(fechaVencimiento.Value.Date - DateTime.Today).TotalDays;

        var vm = new MiSuscripcionIndexViewModel
        {
            NegocioId = baseVm.NegocioId,
            NegocioNombre = baseVm.NegocioNombre,
            RolActual = baseVm.RolActual,
            SedeIdAsignada = baseVm.SedeIdAsignada,
            EsAdministrador = baseVm.EsAdministrador,
            ModuloCodigo = "SUSCRIPCION",
            ModuloNombre = "Mi suscripcion",
            PuedeCrear = baseVm.PuedeCrear,
            PuedeEditar = baseVm.PuedeEditar,
            PuedeEliminar = baseVm.PuedeEliminar,
            Suscripcion = suscripcion,
            ContactoPlataformaEmail = contactoEmail,
            ContactoPlataformaTelefono = contactoTelefono,
            FechaVencimiento = fechaVencimiento,
            DiasParaVencer = diasParaVencer,
            EsModoGratuito = suscripcion?.EsPrueba == true
        };

        return View(vm);
    }
}
