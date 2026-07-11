using System.Text.Json;

namespace FralseTech.Web.ViewModels;

public sealed class LandingPageViewModel
{
    public string PageTitle { get; init; } = string.Empty;
    public string MetaDescription { get; init; } = string.Empty;
    public string CanonicalUrl { get; init; } = string.Empty;
    public string StructuredDataJson { get; init; } = string.Empty;
    public string CommercialName { get; init; } = string.Empty;
    public string LegalName { get; init; } = string.Empty;
    public string Country { get; init; } = string.Empty;
    public string LogoPath { get; init; } = string.Empty;
    public string HeaderTagline { get; init; } = string.Empty;
    public string HeroTitle { get; init; } = string.Empty;
    public string HeroDescription { get; init; } = string.Empty;
    public string HeroPrimaryActionText { get; init; } = string.Empty;
    public string HeroPrimaryActionTarget { get; init; } = string.Empty;
    public string HeroSecondaryActionText { get; init; } = string.Empty;
    public string HeroSecondaryActionTarget { get; init; } = string.Empty;
    public string CallToActionTitle { get; init; } = string.Empty;
    public string CallToActionText { get; init; } = string.Empty;
    public string Email { get; init; } = string.Empty;
    public string MailToUrl { get; init; } = string.Empty;
    public string WhatsAppNumber { get; init; } = string.Empty;
    public string WhatsAppUrl { get; init; } = string.Empty;
    public string WhatsAppMessage { get; init; } = string.Empty;
    public string FacebookUrl { get; init; } = string.Empty;
    public string BusinessHours { get; init; } = string.Empty;
    public string TaxId { get; init; } = string.Empty;
    public int CurrentYear { get; init; }
    public IReadOnlyList<NavigationItem> Navigation { get; init; } = [];
    public IReadOnlyList<HeroPill> HeroPills { get; init; } = [];
    public IReadOnlyList<AboutFeature> AboutFeatures { get; init; } = [];
    public IReadOnlyList<ServiceCard> Services { get; init; } = [];
    public IReadOnlyList<ProductCard> Products { get; init; } = [];
    public IReadOnlyList<AdvantageCard> Advantages { get; init; } = [];
    public IReadOnlyList<ContactCard> ContactCards { get; init; } = [];
}

public sealed record NavigationItem(string Text, string Target);
public sealed record HeroPill(string Text, string IconClass);
public sealed record AboutFeature(string Title, string Description, string IconClass);
public sealed record ServiceCard(string Title, string Description, string IconClass);
public sealed record AdvantageCard(string Title, string Description, string Indicator);
public sealed record ContactCard(string Title, string Value, string Link, string IconClass, string Hint);

public sealed class ProductCard
{
    public string Name { get; init; } = string.Empty;
    public string Category { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public string PrimaryActionText { get; init; } = string.Empty;
    public string PrimaryActionUrl { get; init; } = string.Empty;
    public string SecondaryActionText { get; init; } = string.Empty;
    public string SecondaryActionUrl { get; init; } = string.Empty;
    public string AccentClass { get; init; } = string.Empty;
    public IReadOnlyList<string> Features { get; init; } = [];
    public IReadOnlyList<ProductMockupLine> MockupLines { get; init; } = [];
}

public sealed record ProductMockupLine(string Label, string Value);

public static class FralseTechSiteContent
{
    private const string SiteUrl = "https://fralsetech.com";
    private const string CommercialName = "FRALSE TECH";
    private const string LegalName = "FRALSE TECH S.A.C.";
    private const string Country = "Perú";
    private const string LogoPath = "~/images/logo-fralse-tech.png";
    private const string Email = "contacto@fralsetech.com";
    private const string WhatsAppNumber = "+51 937528701";
    private const string WhatsAppDigits = "51937528701";
    private const string FacebookUrl = "https://facebook.com/";
    private const string FralseContUrl = "https://conta.fralsetech.com/";
    private const string SportsZoneUrl = "https://lazonadeportiva.com";
    private const string BusinessHours = "Lunes a viernes, 9:00 a. m. a 6:00 p. m.";
    private const string TaxId = "20616070186";
    private const string WhatsAppMessage = "Hola, deseo recibir información sobre las soluciones de FRALSE TECH.";
    private const string DemoSubject = "Solicitud de demostración";
    private const string ContactSubject = "Consulta desde fralsetech.com";

    public static LandingPageViewModel BuildLandingPage()
    {
        var whatsappUrl = BuildWhatsAppUrl(WhatsAppMessage);
        var mailToUrl = BuildMailToUrl(ContactSubject);
        var currentYear = DateTime.UtcNow.Year;

        return new LandingPageViewModel
        {
            PageTitle = $"{LegalName} | Desarrollo de software y soluciones empresariales",
            MetaDescription = $"{LegalName} desarrolla sistemas empresariales, plataformas web y soluciones digitales para mejorar la gestión y el crecimiento de las empresas.",
            CanonicalUrl = SiteUrl,
            StructuredDataJson = BuildStructuredDataJson(),
            CommercialName = CommercialName,
            LegalName = LegalName,
            Country = Country,
            LogoPath = LogoPath,
            HeaderTagline = "Soluciones tecnológicas para empresas",
            HeroTitle = "Transformamos ideas en soluciones digitales",
            HeroDescription = "Desarrollamos sistemas empresariales, plataformas web y soluciones digitales diseñadas para mejorar la gestión y el crecimiento de tu negocio.",
            HeroPrimaryActionText = "Conoce nuestros productos",
            HeroPrimaryActionTarget = "#productos",
            HeroSecondaryActionText = "Contáctanos",
            HeroSecondaryActionTarget = "#contacto",
            CallToActionTitle = "¿Tienes una idea o necesitas mejorar la gestión de tu empresa?",
            CallToActionText = "Conversemos sobre cómo una solución tecnológica puede ayudarte a optimizar tus procesos.",
            Email = Email,
            MailToUrl = mailToUrl,
            WhatsAppNumber = WhatsAppNumber,
            WhatsAppUrl = whatsappUrl,
            WhatsAppMessage = WhatsAppMessage,
            FacebookUrl = FacebookUrl,
            BusinessHours = BusinessHours,
            TaxId = TaxId,
            CurrentYear = currentYear,
            Navigation =
            [
                new NavigationItem("Inicio", "#inicio"),
                new NavigationItem("Nosotros", "#nosotros"),
                new NavigationItem("Servicios", "#servicios"),
                new NavigationItem("Productos", "#productos")
            ],
            HeroPills =
            [
                new HeroPill("Sistemas empresariales", "bi bi-building-gear"),
                new HeroPill("Plataformas web", "bi bi-window-stack"),
                new HeroPill("Soluciones en la nube", "bi bi-cloud-check"),
                new HeroPill("Soporte tecnológico", "bi bi-headset")
            ],
            AboutFeatures =
            [
                new AboutFeature("Soluciones a medida", "Diseñamos herramientas alineadas al flujo real de cada operación.", "bi bi-sliders2"),
                new AboutFeature("Tecnología moderna", "Construimos productos actuales, mantenibles y listos para evolucionar.", "bi bi-cpu"),
                new AboutFeature("Enfoque empresarial", "Priorizamos control, visibilidad y orden en los procesos críticos.", "bi bi-briefcase"),
                new AboutFeature("Atención cercana", "Acompañamos cada implementación con comunicación directa y práctica.", "bi bi-people")
            ],
            Services =
            [
                new ServiceCard("Desarrollo de software", "Diseñamos aplicaciones y sistemas adaptados a las necesidades específicas de cada empresa.", "bi bi-code-square"),
                new ServiceCard("Sistemas empresariales", "Creamos herramientas para mejorar la administración, contabilidad, ventas y control operativo.", "bi bi-diagram-3"),
                new ServiceCard("Plataformas web", "Desarrollamos plataformas rápidas, responsivas y disponibles desde cualquier dispositivo.", "bi bi-globe2"),
                new ServiceCard("Soporte y evolución", "Acompañamos el crecimiento de cada proyecto mediante soporte, mejoras y nuevas funcionalidades.", "bi bi-arrow-repeat")
            ],
            Products =
            [
                new ProductCard
                {
                    Name = "FralseCont",
                    Category = "Sistema administrativo y contable",
                    Description = "Plataforma administrativa y contable diseñada para centralizar la información de la empresa, controlar sus operaciones y facilitar la gestión contable.",
                    Status = "Disponible",
                    PrimaryActionText = "Conocer FralseCont",
                    PrimaryActionUrl = FralseContUrl,
                    SecondaryActionText = "Solicitar demostración",
                    SecondaryActionUrl = BuildWhatsAppUrl($"{DemoSubject}: FralseCont"),
                    AccentClass = "product-card--blue",
                    Features =
                    [
                        "Gestión administrativa",
                        "Compras y ventas",
                        "Contabilidad",
                        "Cuentas por cobrar y pagar",
                        "Libros electrónicos",
                        "Reportes empresariales",
                        "Acceso desde la nube"
                    ],
                    MockupLines =
                    [
                        new ProductMockupLine("Panel", "Operaciones centralizadas"),
                        new ProductMockupLine("Módulos", "Ventas, compras y contabilidad"),
                        new ProductMockupLine("Acceso", "Web y nube")
                    ]
                },
                new ProductCard
                {
                    Name = "La Zona Deportiva",
                    Category = "Plataforma para complejos deportivos",
                    Description = "Sistema web para administrar complejos deportivos, organizar horarios, gestionar reservas y ofrecer reservas online a los clientes.",
                    Status = "Disponible",
                    PrimaryActionText = "Visitar plataforma",
                    PrimaryActionUrl = SportsZoneUrl,
                    SecondaryActionText = "Solicitar demostración",
                    SecondaryActionUrl = BuildWhatsAppUrl($"{DemoSubject}: La Zona Deportiva"),
                    AccentClass = "product-card--cyan",
                    Features =
                    [
                        "Reservas online",
                        "Control de horarios",
                        "Gestión de clientes",
                        "Pagos y saldos",
                        "Promociones",
                        "Reportes",
                        "Administración desde celular o computadora"
                    ],
                    MockupLines =
                    [
                        new ProductMockupLine("Agenda", "Reservas por horario"),
                        new ProductMockupLine("Cobranza", "Pagos y saldos"),
                        new ProductMockupLine("Gestión", "Operación móvil y escritorio")
                    ]
                }
            ],
            Advantages =
            [
                new AdvantageCard("Soluciones adaptadas al negocio", "Analizamos procesos reales para diseñar herramientas útiles desde el primer uso.", "Procesos reales"),
                new AdvantageCard("Tecnología moderna y escalable", "Construimos plataformas preparadas para crecer con nuevas áreas, flujos y usuarios.", "Escalabilidad"),
                new AdvantageCard("Acompañamiento cercano", "Mantenemos comunicación directa para ajustar la solución según la operación del cliente.", "Soporte directo"),
                new AdvantageCard("Mejora continua", "Pensamos cada producto como una base evolutiva que puede incorporar nuevas capacidades.", "Evolución")
            ],
            ContactCards =
            [
                new ContactCard("Correo electrónico", Email, mailToUrl, "bi bi-envelope-paper", "Respuesta por correo"),
                new ContactCard("WhatsApp", WhatsAppNumber, whatsappUrl, "bi bi-whatsapp", "Atención comercial")
            ]
        };
    }

    public static string GetDefaultPageTitle() =>
        $"{LegalName} | Desarrollo de software y soluciones empresariales";

    public static string GetDefaultMetaDescription() =>
        $"{LegalName} desarrolla sistemas empresariales, plataformas web y soluciones digitales para mejorar la gestión y el crecimiento de las empresas.";

    public static string GetDefaultCanonicalUrl() => SiteUrl;

    public static string GetStructuredDataJson() => BuildStructuredDataJson();

    private static string BuildStructuredDataJson()
    {
        var payload = new
        {
            @context = "https://schema.org",
            @type = "Organization",
            name = LegalName,
            alternateName = CommercialName,
            url = SiteUrl,
            logo = $"{SiteUrl}/images/logo-fralse-tech.png",
            email = Email,
            areaServed = Country,
            sameAs = new[] { FacebookUrl },
            contactPoint = new[]
            {
                new
                {
                    @type = "ContactPoint",
                    contactType = "sales",
                    areaServed = Country,
                    availableLanguage = "es",
                    email = Email,
                    telephone = WhatsAppNumber
                }
            }
        };

        return JsonSerializer.Serialize(payload);
    }

    private static string BuildWhatsAppUrl(string message) =>
        $"https://wa.me/{WhatsAppDigits}?text={Uri.EscapeDataString(message)}";

    private static string BuildMailToUrl(string subject) =>
        $"mailto:{Email}?subject={Uri.EscapeDataString(subject)}";
}
