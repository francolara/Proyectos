using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SistemaAdministrativoWeb.Configuration;
using SistemaAdministrativoWeb.ViewModels.Legal;

namespace SistemaAdministrativoWeb.Controllers;

[AllowAnonymous]
public class LegalController(
    IOptions<BusinessInformationOptions> businessInformationOptions,
    IOptions<LegalDocumentsOptions> legalDocumentsOptions) : Controller
{
    // Firma: FRANCO LARA - 10/07/2026 | Agrega las rutas publicas legales de FralseCont con contenido configurable, metadatos basicos y aviso preliminar editable desde configuracion.

    [HttpGet("/terminos-y-condiciones")]
    public IActionResult Terminos()
    {
        return View(CrearModelo(
            "Terminos y condiciones | FralseCont",
            "Consulta los terminos y condiciones de uso de FralseCont, plataforma contable operada por FRALSE TECH S.A.C.",
            legalDocumentsOptions.Value.TermsLastUpdated));
    }

    [HttpGet("/politica-de-privacidad")]
    public IActionResult Privacidad()
    {
        return View(CrearModelo(
            "Politica de privacidad | FralseCont",
            "Conoce como FralseCont recopila, utiliza, almacena y protege la informacion de sus usuarios y empresas registradas.",
            legalDocumentsOptions.Value.PrivacyLastUpdated));
    }

    [HttpGet("/cancelaciones-y-devoluciones")]
    public IActionResult Cancelaciones()
    {
        return View(CrearModelo(
            "Politica de cancelaciones y devoluciones | FralseCont",
            "Consulta las condiciones de cancelacion, renovacion y evaluacion de devoluciones aplicables a los planes de FralseCont.",
            legalDocumentsOptions.Value.CancellationsLastUpdated));
    }

    private LegalDocumentViewModel CrearModelo(string title, string metaDescription, string lastUpdated)
    {
        return new LegalDocumentViewModel
        {
            Title = title,
            MetaDescription = metaDescription,
            LastUpdated = lastUpdated,
            ShowDraftNotice = businessInformationOptions.Value.LegalDocumentsDraftNotice,
            BusinessInformation = businessInformationOptions.Value
        };
    }
}
