namespace SistemaAdministrativoWeb.ViewModels.Ayuda;

public sealed class AyudaModuloViewModel
{
    public string Clave { get; init; } = string.Empty;
    public string Titulo { get; init; } = string.Empty;
    public string Icono { get; init; } = "bi-question-circle";
    public string Resumen { get; init; } = string.Empty;
    public IReadOnlyCollection<AyudaPreguntaViewModel> Preguntas { get; init; } = [];
}
