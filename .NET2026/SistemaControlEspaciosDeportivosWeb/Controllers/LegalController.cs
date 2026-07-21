using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SistemaControlEspaciosDeportivosWeb.Configuration;
using SistemaControlEspaciosDeportivosWeb.ViewModels.Legal;

namespace SistemaControlEspaciosDeportivosWeb.Controllers;

[AllowAnonymous]
public class LegalController(
    IOptions<BusinessInformationOptions> businessInformationOptions,
    IOptions<LegalDocumentsOptions> legalDocumentsOptions) : Controller
{
    // Firma: FRANCO LARA - 20/07/2026 | Agrega documentos legales publicos para los planes de La Zona Deportiva.

    [HttpGet("/terminos-y-condiciones")]
    public IActionResult Terminos() => View(CrearModelo("Términos y condiciones", legalDocumentsOptions.Value.TermsLastUpdated));

    [HttpGet("/politica-de-privacidad")]
    public IActionResult Privacidad() => View(CrearModelo("Política de privacidad", legalDocumentsOptions.Value.PrivacyLastUpdated));

    [HttpGet("/cancelaciones-y-devoluciones")]
    public IActionResult Cancelaciones() => View(CrearModelo("Cancelaciones y devoluciones", legalDocumentsOptions.Value.CancellationsLastUpdated));

    private LegalDocumentViewModel CrearModelo(string title, string lastUpdated) => new()
    {
        Title = title,
        LastUpdated = lastUpdated,
        ShowDraftNotice = businessInformationOptions.Value.LegalDocumentsDraftNotice,
        BusinessInformation = businessInformationOptions.Value
    };
}
