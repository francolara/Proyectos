namespace SistemaControlEspaciosDeportivosWeb.Configuration;

public sealed class BusinessInformationOptions
{
    public const string SectionName = "BusinessInformation";

    public string CommercialName { get; set; } = "La Zona Deportiva";
    public string LegalName { get; set; } = "FRALSE TECH S.A.C.";
    public string Ruc { get; set; } = "20616070186";
    public string Address { get; set; } = "Quimper 459, Villa María del Triunfo, Lima, Perú";
    public string SupportEmail { get; set; } = "informes@lazonadeportiva.com";
    public string SupportPhone { get; set; } = "+51 937528701";
    public string WhatsAppDisplay { get; set; } = "+51 937 528 701";
    public string WhatsAppNumber { get; set; } = "51937528701";
    public string WhatsAppMessage { get; set; } = "Hola, deseo recibir información sobre La Zona Deportiva.";
    public string SupportHours { get; set; } = "Lunes a viernes de 9:00 a 18:00 horas";
    public bool LegalDocumentsDraftNotice { get; set; }
}
