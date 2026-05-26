namespace SistemaControlEspaciosDeportivosWeb.ViewModels;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;

public class OnboardingChecklistViewModel
{
    public int NegocioId { get; set; }
    public bool ConfigNombreComercialOk { get; set; }
    public bool ConfigTipoDocumentoOk { get; set; }
    public bool ConfigMonedaOk { get; set; }
    public bool ConfigCpeCondicionesOk { get; set; }
    public bool MaestroTipoDeporteOk { get; set; }
    public bool MaestroTipoSueloOk { get; set; }
    public bool MaestroFormaPagoOk { get; set; }
    public bool MaestroMonedaOk { get; set; }
    public bool MaestroTipoDocumentoOk { get; set; }
    public bool MaestroSerieDocumentoOk { get; set; }
    public bool SedeMinimaOk { get; set; }
    public bool EspacioMinimoOk { get; set; }
    public bool ChecklistCompleto { get; set; }
}

public class OnboardingPasoItemViewModel
{
    public byte Paso { get; set; }
    public string Titulo { get; set; } = string.Empty;
    public string Descripcion { get; set; } = string.Empty;
    public bool Completado { get; set; }
    public bool EsActual { get; set; }
    public string UrlAccion { get; set; } = string.Empty;
}

public class OnboardingDashboardViewModel
{
    public int NegocioId { get; set; }
    public string NegocioNombre { get; set; } = string.Empty;
    public string? LogoUrl { get; set; }
    public byte PasoActual { get; set; } = 1;
    public int PasosCompletados { get; set; }
    public int TotalPasos { get; set; } = 5;
    public bool ChecklistCompleto { get; set; }
    public OnboardingChecklistViewModel Checklist { get; set; } = new();
    public List<OnboardingPasoItemViewModel> Pasos { get; set; } = new();
    public OnboardingConfiguracionFormViewModel ConfiguracionForm { get; set; } = new();
    public OnboardingMaestrosFormViewModel MaestrosForm { get; set; } = new();
    public OnboardingSedeFormViewModel SedeForm { get; set; } = new();
    public OnboardingEspacioFormViewModel EspacioForm { get; set; } = new();
}

public class OnboardingConfiguracionFormViewModel
{
    public int NegocioId { get; set; }
    [Required] public string NombreComercial { get; set; } = string.Empty;
    [Required] public string TipoDocumento { get; set; } = "1";
    public string? NumeroDocumento { get; set; }
    [Range(1, int.MaxValue)] public int MonedaId { get; set; } = 1;
    public List<SelectListItem> TiposDocumento { get; set; } = new();
    public List<SelectListItem> Monedas { get; set; } = new();
}

public class OnboardingMaestrosFormViewModel
{
    public int NegocioId { get; set; }
    public int? TipoDeporteSuperId { get; set; }
    public int? TipoSueloSuperId { get; set; }
    public int? MonedaSuperId { get; set; }
    public string? FormaPagoNombre { get; set; }
    public string? CodigoSunatDocumento { get; set; }
    public string? SerieDocumento { get; set; }
    public List<SelectListItem> TiposDeporteSuper { get; set; } = new();
    public List<SelectListItem> TiposSueloSuper { get; set; } = new();
    public List<SelectListItem> MonedasSuper { get; set; } = new();
    public List<SelectListItem> TiposDocumentoComprobanteSuper { get; set; } = new();
}

public class OnboardingSedeFormViewModel
{
    public int NegocioId { get; set; }
    [Required] public string Nombre { get; set; } = string.Empty;
    [Required] public string Direccion { get; set; } = string.Empty;
    [Required] public string CodigoUbigeo { get; set; } = string.Empty;
    public TimeOnly HoraApertura { get; set; } = new(8, 0);
    public TimeOnly HoraCierre { get; set; } = new(23, 0);
    public List<int> ServiciosSeleccionados { get; set; } = new();
    public List<SelectListItem> ServiciosDisponibles { get; set; } = new();
}

public class OnboardingEspacioFormViewModel
{
    public int NegocioId { get; set; }
    [Range(1, int.MaxValue)] public int SedeId { get; set; }
    [Range(1, int.MaxValue)] public int TipoDeporteId { get; set; }
    [Range(1, int.MaxValue)] public int TipoSueloId { get; set; }
    [Required] public string Codigo { get; set; } = string.Empty;
    [Required] public string Nombre { get; set; } = string.Empty;
    [Range(1, 200)] public int Capacidad { get; set; } = 10;
    [Range(0.01, 999999)] public decimal PrecioBase { get; set; } = 1m;
    public List<SelectListItem> Sedes { get; set; } = new();
    public List<SelectListItem> TiposDeporte { get; set; } = new();
    public List<SelectListItem> TiposSuelo { get; set; } = new();
}
