namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class RegistroCuentaAdministradoraConEmpresaResult
{
    public int IdCuentaAdministradora { get; init; }
    public int IdEmpresa { get; init; }
    public DateOnly? FechaInicioPrueba { get; init; }
    public DateOnly? FechaFinPrueba { get; init; }
}
