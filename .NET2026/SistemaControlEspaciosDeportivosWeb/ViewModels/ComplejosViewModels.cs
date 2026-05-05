namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public class ComplejosIndexViewModel
{
    public List<ComplejoListadoItemViewModel> Complejos { get; set; } = new();
}

public class ComplejoListadoItemViewModel
{
    public int NegocioId { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Slug { get; set; } = string.Empty;
    public int TotalSedes { get; set; }
    public string? DireccionReferencia { get; set; }
}

public class ComplejoDetalleViewModel
{
    public int NegocioId { get; set; }
    public string Nombre { get; set; } = string.Empty;
    public string Slug { get; set; } = string.Empty;
    public List<SedePublicaViewModel> Sedes { get; set; } = new();
}
