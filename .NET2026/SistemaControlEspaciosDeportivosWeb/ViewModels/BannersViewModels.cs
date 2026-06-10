using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Http;

namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public enum BannerTipo
{
    Home = 1,
    Login = 2,
    Registro = 3
}

public class BannersAdminIndexViewModel : ModuloBaseViewModel
{
    public bool? SoloActivos { get; set; }
    public List<BannerAdminItemViewModel> Banners { get; set; } = new();
    public BannerAdminFormViewModel Form { get; set; } = new();
    public string? MensajeUi { get; set; }
    public string? Error { get; set; }
}

public class BannerAdminItemViewModel
{
    public int Id { get; set; }
    public string Titulo { get; set; } = string.Empty;
    public string? Subtitulo { get; set; }
    public string? Descripcion { get; set; }
    public string? BotonTexto { get; set; }
    public string? BotonUrl { get; set; }
    public string ImagenUrl { get; set; } = string.Empty;
    public string? ImagenUrlMobile { get; set; }
    public int TipoBanner { get; set; } = (int)BannerTipo.Home;
    public string TipoBannerNombre { get; set; } = "Home";
    public int Orden { get; set; }
    public bool Activo { get; set; }
    public DateOnly? FechaInicio { get; set; }
    public DateOnly? FechaFin { get; set; }
}

public class BannerAdminFormViewModel
{
    public int? Id { get; set; }

    [StringLength(120)]
    public string? Titulo { get; set; }

    [StringLength(220)]
    public string? Subtitulo { get; set; }

    [StringLength(400)]
    public string? Descripcion { get; set; }

    [StringLength(40)]
    public string? BotonTexto { get; set; }

    [StringLength(300)]
    public string? BotonUrl { get; set; }

    [StringLength(500)]
    public string? ImagenUrl { get; set; }
    [StringLength(500)]
    public string? ImagenUrlMobile { get; set; }

    public IFormFile? ImagenArchivo { get; set; }
    public IFormFile? ImagenArchivoMobile { get; set; }

    [Range(1, 3, ErrorMessage = "Selecciona un tipo de banner valido.")]
    public int TipoBanner { get; set; } = (int)BannerTipo.Home;

    [Range(1, 999)]
    public int Orden { get; set; } = 1;

    public bool Activo { get; set; } = true;
    public DateOnly? FechaInicio { get; set; }
    public DateOnly? FechaFin { get; set; }
}
