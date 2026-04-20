namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class PlataformaIndexViewModel
{
    public string CorreoUsuario { get; set; } = string.Empty;
    public int TotalBanners { get; set; }
    public int BannersActivos { get; set; }
    public int BannersInactivos { get; set; }
}

public class PlataformaPortalConfigViewModel
{
    public string? BeneficiosTitulo { get; set; } = "Todo lo que necesitas para gestionar tus canchas deportivas";
    public string? BeneficiosSubtitulo { get; set; } = "SportCenter integra reservas, sedes, pagos y reportes en una sola plataforma para crecer tu operacion.";
    public string? Beneficio1Titulo { get; set; } = "Sistema de reservas";
    public string? Beneficio1Detalle { get; set; } = "Controla la disponibilidad por horario con agenda visual y registro de clientes en segundos.";
    public string? Beneficio2Titulo { get; set; } = "Multiples sedes";
    public string? Beneficio2Detalle { get; set; } = "Administra distintos complejos deportivos desde un solo panel operativo.";
    public string? Beneficio3Titulo { get; set; } = "Pagos seguros";
    public string? Beneficio3Detalle { get; set; } = "Gestiona adelantos, saldos y comprobantes con trazabilidad por reserva.";
    public string? Beneficio4Titulo { get; set; } = "Promociones especiales";
    public string? Beneficio4Detalle { get; set; } = "Crea descuentos por sede, dia y horario para mejorar ocupacion en horas clave.";
    public string? Beneficio5Titulo { get; set; } = "Estadisticas detalladas";
    public string? Beneficio5Detalle { get; set; } = "Analiza ingresos, ocupacion y rendimiento para tomar decisiones con datos.";
    public string? Beneficio6Titulo { get; set; } = "Mayor visibilidad";
    public string? Beneficio6Detalle { get; set; } = "Publica tu negocio en el portal y recibe solicitudes online de nuevos clientes.";

    public string? CtaTitulo { get; set; } = "Unete a la comunidad de SportCenter";
    public string? CtaSubtitulo { get; set; } = "Registra tu club deportivo y comienza a gestionar tus canchas de manera eficiente.";
    public string? CtaBotonClubTexto { get; set; } = "Registrar mi club";
    public string? CtaBotonClubUrl { get; set; } = "/Home/SoftwareClubes";
    public string? CtaBotonUsuarioTexto { get; set; } = "Crear cuenta personal";
    public string? CtaBotonUsuarioUrl { get; set; } = "/Identity/Account/Register";
    public string? MarcaTitulo { get; set; } = "SportCenter";
    public string? MarcaDescripcion { get; set; } = "La plataforma lider para la reserva y gestion de canchas deportivas.";
    public string? ContactoEmail { get; set; } = "contacto@sportcenter.com";
    public string? ContactoTelefono { get; set; } = "+51 900 000 000";
    public string? SiguenosFacebookUrl { get; set; } = string.Empty;
    public string? SiguenosInstagramUrl { get; set; } = string.Empty;
    public string? SiguenosWhatsappUrl { get; set; } = string.Empty;
}

public class PlataformaNegociosAdminViewModel
{
    public string? Buscar { get; set; }
    public string EstadoContrato { get; set; } = "todos";
    public int Pagina { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas { get; set; } = 1;
    public List<PlataformaNegocioLimiteItemViewModel> Negocios { get; set; } = new();
}

public class PlataformaNegocioLimiteItemViewModel
{
    public int NegocioId { get; set; }
    public string NombreComercial { get; set; } = string.Empty;
    public bool Activo { get; set; }
    public int SedesPermitidas { get; set; } = 2;
    public int EspaciosPermitidos { get; set; } = 6;
    public int UsuariosPermitidos { get; set; } = 3;
    public int EstadoSuscripcion { get; set; }
    public string EstadoSuscripcionNombre { get; set; } = "Sin suscripcion";
    public bool EsPrueba { get; set; }
    public DateTime? FechaInicioPrueba { get; set; }
    public DateTime? FechaFinPrueba { get; set; }
    public string? TipoCobro { get; set; }
    public DateTime? FechaInicioPlan { get; set; }
    public DateTime? FechaFinPlan { get; set; }
    public int DiasGracia { get; set; } = 5;
    public DateTime? FechaFinGracia { get; set; }
}

public class PlataformaAltasClubesAdminViewModel
{
    public int? Estado { get; set; }
    public int DiasPruebaDefault { get; set; } = 30;
    public int Pagina { get; set; } = 1;
    public int TamanoPagina { get; set; } = 20;
    public int TotalRegistros { get; set; }
    public int TotalPaginas { get; set; } = 1;
    public int TotalPendientes { get; set; }
    public int TotalAprobados { get; set; }
    public int TotalRechazados { get; set; }
    public List<AltaClubItemViewModel> Solicitudes { get; set; } = new();
}

public class MiSuscripcionNegocioViewModel
{
    public int NegocioId { get; set; }
    public int EstadoSuscripcion { get; set; }
    public string EstadoSuscripcionNombre { get; set; } = "Pendiente de activacion";
    public bool EsPrueba { get; set; }
    public DateTime? FechaInicioPrueba { get; set; }
    public DateTime? FechaFinPrueba { get; set; }
    public string? TipoCobro { get; set; }
    public DateTime? FechaInicioPlan { get; set; }
    public DateTime? FechaFinPlan { get; set; }
    public int DiasGracia { get; set; } = 5;
    public DateTime? FechaFinGracia { get; set; }
}

public class MiSuscripcionIndexViewModel : ModuloBaseViewModel
{
    public MiSuscripcionNegocioViewModel? Suscripcion { get; set; }
    public string ContactoPlataformaEmail { get; set; } = string.Empty;
    public string ContactoPlataformaTelefono { get; set; } = string.Empty;
    public DateTime? FechaVencimiento { get; set; }
    public int? DiasParaVencer { get; set; }
    public bool EsModoGratuito { get; set; }
}
