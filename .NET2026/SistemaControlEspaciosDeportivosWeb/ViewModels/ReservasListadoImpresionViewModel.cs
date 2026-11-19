namespace SistemaControlEspaciosDeportivosWeb.ViewModels;

public sealed record ReservasListadoImpresionViewModel(
    string NegocioNombre,
    DateOnly FechaDesde,
    DateOnly FechaHasta,
    string Sede,
    string Espacio,
    string Estados,
    IReadOnlyList<ReservaItemViewModel> Reservas,
    ReservasListadoResumenViewModel Resumen);
