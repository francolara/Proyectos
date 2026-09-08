namespace SistemaControlEspaciosDeportivosWeb.ViewModels.Ayuda;

public sealed class AyudaIndexViewModel
{
    public ModuloBaseViewModel Base { get; init; } = new();
    public string CategoriaSeleccionadaClave { get; init; } = string.Empty;
    public string ModuloSeleccionadoClave { get; init; } = string.Empty;
    public string ModuloSeleccionadoTitulo { get; init; } = string.Empty;
    public string? ModuloSolicitado { get; init; }
    public int TotalPreguntas { get; init; }
    public IReadOnlyCollection<AyudaCategoriaViewModel> Categorias { get; init; } = [];
}

public sealed class AyudaCategoriaViewModel
{
    public string Clave { get; init; } = string.Empty;
    public string Titulo { get; init; } = string.Empty;
    public string Icono { get; init; } = "bi-collection";
    public string Descripcion { get; init; } = string.Empty;
    public IReadOnlyCollection<AyudaModuloViewModel> Modulos { get; init; } = [];
}

public sealed class AyudaModuloViewModel
{
    public string Clave { get; init; } = string.Empty;
    public string Titulo { get; init; } = string.Empty;
    public string Icono { get; init; } = "bi-question-circle";
    public string Resumen { get; init; } = string.Empty;
    public IReadOnlyCollection<AyudaPreguntaViewModel> Preguntas { get; init; } = [];
}

public sealed class AyudaPreguntaViewModel
{
    public string Id { get; init; } = string.Empty;
    public string Pregunta { get; init; } = string.Empty;
    public string Respuesta { get; init; } = string.Empty;
}
