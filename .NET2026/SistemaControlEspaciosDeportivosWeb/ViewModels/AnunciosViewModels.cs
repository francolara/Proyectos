using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class PopupPromocionPublicoViewModel
{
    public const string OrientacionVertical = "V";
    public const string OrientacionHorizontal = "H";

    public int IdPopupPromocion { get; set; }
    public string Titulo { get; set; } = string.Empty;
    public string? Subtitulo { get; set; }
    public string? Descripcion { get; set; }
    public string ImagenUrl { get; set; } = string.Empty;
    public string? TextoBoton { get; set; }
    public string? UrlBoton { get; set; }
    public string? UrlImagen { get; set; }
    public int Orden { get; set; }
    public bool AbrirNuevaPestana { get; set; }
    public string Orientacion { get; set; } = OrientacionVertical;
    public bool EsHorizontal => string.Equals(Orientacion, OrientacionHorizontal, StringComparison.OrdinalIgnoreCase);
    public bool EsVertical => !EsHorizontal;
}

public class PopupPromocionAdminItemViewModel : PopupPromocionPublicoViewModel
{
    public bool Activo { get; set; }
    public DateOnly? FechaInicio { get; set; }
    public DateOnly? FechaFin { get; set; }
    public DateTime FechaCreacion { get; set; }
    public DateTime? FechaModificacion { get; set; }
}

public class PopupPromocionAdminFormViewModel
{
    public int? IdPopupPromocion { get; set; }

    [Display(Name = "Titulo")]
    [Required(ErrorMessage = "Ingresa un titulo.")]
    [StringLength(120, ErrorMessage = "El titulo no puede superar los 120 caracteres.")]
    public string Titulo { get; set; } = string.Empty;

    [Display(Name = "Descripcion corta")]
    [StringLength(260, ErrorMessage = "La descripcion no puede superar los 260 caracteres.")]
    public string? Descripcion { get; set; }

    [Display(Name = "Subtitulo")]
    [StringLength(140, ErrorMessage = "El subtitulo no puede superar los 140 caracteres.")]
    public string? Subtitulo { get; set; }

    [Display(Name = "Imagen actual")]
    [StringLength(500)]
    public string? ImagenUrl { get; set; }

    [Display(Name = "Imagen")]
    public IFormFile? ImagenArchivo { get; set; }

    [Display(Name = "Orientacion")]
    [Required(ErrorMessage = "Selecciona la orientacion del anuncio.")]
    [StringLength(1)]
    public string Orientacion { get; set; } = PopupPromocionPublicoViewModel.OrientacionVertical;

    [Display(Name = "Texto del boton")]
    [StringLength(40, ErrorMessage = "El texto del boton no puede superar los 40 caracteres.")]
    public string? TextoBoton { get; set; }

    [Display(Name = "URL del boton")]
    [StringLength(300, ErrorMessage = "La URL del boton no puede superar los 300 caracteres.")]
    public string? UrlBoton { get; set; }

    [Display(Name = "URL de la imagen")]
    [StringLength(300, ErrorMessage = "La URL de la imagen no puede superar los 300 caracteres.")]
    public string? UrlImagen { get; set; }

    [Display(Name = "Orden")]
    [Range(1, 999, ErrorMessage = "El orden debe ser mayor o igual a 1.")]
    public int Orden { get; set; } = 1;

    [Display(Name = "Activo")]
    public bool Activo { get; set; } = true;

    [Display(Name = "Fecha inicio")]
    public DateOnly? FechaInicio { get; set; }

    [Display(Name = "Fecha fin")]
    public DateOnly? FechaFin { get; set; }

    [Display(Name = "Abrir en nueva pestana")]
    public bool AbrirNuevaPestana { get; set; } = true;
}

public class PopupPromocionConfigViewModel
{
    [Display(Name = "Activar popup automatico")]
    public bool ActivarPopupAutomatico { get; set; } = true;

    [Display(Name = "Segundos de espera antes de mostrar")]
    [Range(0, 30, ErrorMessage = "Los segundos de espera deben estar entre 0 y 30.")]
    public int SegundosEsperaAntesDeMostrar { get; set; } = 1;

    [Display(Name = "Activar autoplay slider")]
    public bool ActivarAutoplaySlider { get; set; } = true;

    [Display(Name = "Velocidad de autoplay (ms)")]
    [Range(1000, 20000, ErrorMessage = "La velocidad de autoplay debe estar entre 1000 y 20000 ms.")]
    public int VelocidadAutoplayMs { get; set; } = 4500;

    [Display(Name = "Mostrar flechas")]
    public bool MostrarFlechas { get; set; } = true;

    [Display(Name = "Mostrar indicadores")]
    public bool MostrarIndicadores { get; set; } = true;
}

public class PopupPromocionesPublicasPartialViewModel
{
    public List<PopupPromocionPublicoViewModel> PopupPromociones { get; set; } = new();
    public PopupPromocionConfigViewModel Config { get; set; } = new();
    public PlataformaPortalConfigViewModel PortalConfig { get; set; } = new();
}

public class PopupPromocionesAdminIndexViewModel : ModuloBaseViewModel
{
    public List<PopupPromocionAdminItemViewModel> Anuncios { get; set; } = new();
    public PopupPromocionAdminFormViewModel Form { get; set; } = new();
    public PopupPromocionConfigViewModel Config { get; set; } = new();
    public string? MensajeUi { get; set; }
    public string? Error { get; set; }
}
