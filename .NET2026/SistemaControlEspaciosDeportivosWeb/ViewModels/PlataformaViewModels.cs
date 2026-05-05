using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class PlataformaIndexViewModel
{
    public string CorreoUsuario { get; set; } = string.Empty;
    public int TotalBanners { get; set; }
    public int BannersActivos { get; set; }
    public int BannersInactivos { get; set; }
    public int TotalNegocios { get; set; }
    public int NegociosConContrato { get; set; }
    public int NegociosEnPrueba { get; set; }
    public int NegociosVencidos { get; set; }
    public int TotalSolicitudesPendientes { get; set; }
    public int TotalSolicitudesAprobadas { get; set; }
    public int TotalSolicitudesRechazadas { get; set; }
    public int TotalReferencialesActivos { get; set; }
    public int TotalReferencialesInactivos { get; set; }
    public int TotalAnuncios { get; set; }
    public int AnunciosActivos { get; set; }
    public int AnunciosInactivos { get; set; }
    public int AnunciosVigentesHoy { get; set; }
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

    [EmailAddress(ErrorMessage = "Ingresa un correo valido en Notificacion 1.")]
    [StringLength(500, ErrorMessage = "Notificacion 1 permite hasta 500 caracteres.")]
    public string? NotificacionCorreo1 { get; set; } = string.Empty;

    [EmailAddress(ErrorMessage = "Ingresa un correo valido en Notificacion 2.")]
    [StringLength(500, ErrorMessage = "Notificacion 2 permite hasta 500 caracteres.")]
    public string? NotificacionCorreo2 { get; set; } = string.Empty;
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
    public string? CorreoContacto { get; set; }
    public string? TelefonoContacto { get; set; }
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
    public int SedesPermitidas { get; set; } = 2;
    public int EspaciosPermitidos { get; set; } = 6;
    public int UsuariosPermitidos { get; set; } = 3;
}

public class PlataformaReferencialesExternosViewModel
{
    [Required(ErrorMessage = "Debes seleccionar un departamento.")]
    [StringLength(2)]
    public string CodigoDepartamento { get; set; } = string.Empty;

    [Required(ErrorMessage = "Debes seleccionar una provincia.")]
    [StringLength(4)]
    public string CodigoProvincia { get; set; } = string.Empty;

    [Required(ErrorMessage = "Debes seleccionar un distrito.")]
    [StringLength(6)]
    public string CodigoUbigeo { get; set; } = string.Empty;

    [Required(ErrorMessage = "Debes seleccionar un tipo de deporte.")]
    [Range(1, int.MaxValue, ErrorMessage = "Tipo de deporte invalido.")]
    public int TipoDeporteSuperId { get; set; }

    [Required(ErrorMessage = "Debes ingresar una palabra clave.")]
    [StringLength(120, ErrorMessage = "Palabra clave excede la longitud permitida.")]
    public string PalabraClave { get; set; } = string.Empty;

    [Range(1, 60, ErrorMessage = "El maximo de resultados permitidos es 60.")]
    public int MaxResultados { get; set; } = 40;
    public bool DescargarTelefonos { get; set; } = true;
    public bool DescargarFotos { get; set; } = false;

    public List<SelectListItem> DepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> ProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> DistritosUbigeo { get; set; } = new();
    public List<SelectListItem> TiposDeporte { get; set; } = new();
    public ReferencialesExternosSyncResultadoViewModel? Resultado { get; set; }

    [StringLength(180, ErrorMessage = "La busqueda por nombre permite hasta 180 caracteres.")]
    public string? BuscarNombre { get; set; }

    [StringLength(2)]
    public string? FiltroCodigoDepartamento { get; set; }

    [StringLength(4)]
    public string? FiltroCodigoProvincia { get; set; }

    [StringLength(6)]
    public string? FiltroCodigoUbigeo { get; set; }

    public bool IncluirInactivos { get; set; }
    public int PaginaListado { get; set; } = 1;
    public int TamanoPaginaListado { get; set; } = 20;
    public int TotalRegistrosListado { get; set; }
    public int TotalPaginasListado { get; set; } = 1;
    public List<SelectListItem> FiltroDepartamentosUbigeo { get; set; } = new();
    public List<SelectListItem> FiltroProvinciasUbigeo { get; set; } = new();
    public List<SelectListItem> FiltroDistritosUbigeo { get; set; } = new();
    public List<ReferencialExternoAdminItemViewModel> ReferencialesListado { get; set; } = new();
    public bool BarridoHabilitado { get; set; }
}

public class ReferencialesExternosSyncResultadoViewModel
{
    public int Procesados { get; set; }
    public int Insertados { get; set; }
    public int Actualizados { get; set; }
    public int Omitidos { get; set; }
    public int Errores { get; set; }
    public List<string> Detalles { get; set; } = new();
}

public class ReferencialExternoAdminItemViewModel
{
    public int Id { get; set; }
    public string NombreComplejo { get; set; } = string.Empty;
    public string? NombreEspacio { get; set; }
    public string CodigoUbigeo { get; set; } = string.Empty;
    public int TipoDeporteSuperId { get; set; }
    public string TipoDeporte { get; set; } = string.Empty;
    public string Departamento { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
    public string Distrito { get; set; } = string.Empty;
    public string? Direccion { get; set; }
    public string? TelefonoContacto { get; set; }
    public string? GoogleMapsUrl { get; set; }
    public bool Activo { get; set; }
    public DateTime? FechaActualizacion { get; set; }
    public string? UsuarioActualizacion { get; set; }
}
