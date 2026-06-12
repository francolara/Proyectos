using Microsoft.AspNetCore.Mvc;
using SistemaControlEspaciosDeportivosWeb.Services;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.ViewComponents;

public class PopupPromocionesPublicasViewComponent(ISportCenterStoredProcedureService spService) : ViewComponent
{
    public async Task<IViewComponentResult> InvokeAsync()
    {
        var popupPromocionesTask = spService.HomeListarPopupPromocionesActivasAsync();
        var popupConfigTask = CargarPopupPromocionesConfigAsync();
        var portalConfigTask = CargarPortalConfigAsync();

        await Task.WhenAll(popupPromocionesTask, popupConfigTask, portalConfigTask);

        var vm = new PopupPromocionesPublicasPartialViewModel
        {
            PopupPromociones = popupPromocionesTask.Result,
            Config = popupConfigTask.Result,
            PortalConfig = portalConfigTask.Result
        };

        return View("~/Views/Shared/_PopupPromocionesPublicas.cshtml", vm);
    }

    private async Task<PopupPromocionConfigViewModel> CargarPopupPromocionesConfigAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);

        return new PopupPromocionConfigViewModel
        {
            ActivarPopupAutomatico = LeerBool(await Get("POPUP_PROMO_AUTO_ENABLED"), true),
            SegundosEsperaAntesDeMostrar = LeerEntero(await Get("POPUP_PROMO_DELAY_SECONDS"), 1, 0, 30),
            ActivarAutoplaySlider = LeerBool(await Get("POPUP_PROMO_AUTOPLAY_ENABLED"), true),
            VelocidadAutoplayMs = LeerEntero(await Get("POPUP_PROMO_AUTOPLAY_MS"), 4500, 1000, 20000),
            MostrarFlechas = LeerBool(await Get("POPUP_PROMO_SHOW_ARROWS"), true),
            MostrarIndicadores = LeerBool(await Get("POPUP_PROMO_SHOW_INDICATORS"), true)
        };
    }

    private async Task<PlataformaPortalConfigViewModel> CargarPortalConfigAsync()
    {
        async Task<string?> Get(string key) => await spService.ParametrosGlobalesObtenerValorAsync(key);

        return new PlataformaPortalConfigViewModel
        {
            SiguenosWhatsappUrl = (await Get("HOME_PORTAL_WHATSAPP_URL")) ?? string.Empty
        };
    }

    private static bool LeerBool(string? valor, bool fallback)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return fallback;

        valor = valor.Trim();
        if (valor == "1")
            return true;
        if (valor == "0")
            return false;

        return bool.TryParse(valor, out var parsed) ? parsed : fallback;
    }

    private static int LeerEntero(string? valor, int fallback, int min, int max)
    {
        if (!int.TryParse((valor ?? string.Empty).Trim(), out var parsed))
            return fallback;

        return Math.Clamp(parsed, min, max);
    }
}
