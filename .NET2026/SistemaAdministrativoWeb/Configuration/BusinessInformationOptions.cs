namespace SistemaAdministrativoWeb.Configuration;

public sealed class BusinessInformationOptions
{
    public const string SectionName = "BusinessInformation";

    public string CommercialName { get; set; } = "FralseCont";
    public string Slogan { get; set; } = "Tu contabilidad en la nube, simple y segura.";
    public string LegalName { get; set; } = "FRALSE TECH S.A.C.";
    public string Ruc { get; set; } = "20616070186";
    public string Address { get; set; } = "Quimper 459, Villa María del Triunfo";
    public string City { get; set; } = "Lima";
    public string Country { get; set; } = "Perú";
    public string SupportEmail { get; set; } = "contacto@fralsetech.com";
    public string PrivacyEmail { get; set; } = "contacto@fralsetech.com";
    public string WhatsAppDisplay { get; set; } = "+51 937 528 701";
    public string WhatsAppNumber { get; set; } = "51937528701";
    public string WhatsAppMessage { get; set; } = "Hola, estoy interesado en FralseCont. Deseo recibir información sobre los planes y funcionalidades.";
    public string SupportHours { get; set; } = "Lunes a viernes de 9:00 a 18:00 horas";
    public bool LegalDocumentsDraftNotice { get; set; } = true;

    public string FullAddress =>
        string.Join(", ", new[] { Address, City, Country }.Where(x => !string.IsNullOrWhiteSpace(x)));
}
