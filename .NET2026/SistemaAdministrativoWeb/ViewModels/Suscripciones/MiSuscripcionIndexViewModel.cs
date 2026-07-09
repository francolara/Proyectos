namespace SistemaAdministrativoWeb.ViewModels.Suscripciones;

public sealed class MiSuscripcionIndexViewModel
{
    public int IdEmpresa { get; init; }
    public string NombreEmpresa { get; init; } = "Empresa";
    public string? Ruc { get; init; }
    public int IdCuentaAdministradora { get; init; }
    public string CodigoCuenta { get; init; } = string.Empty;
    public string NombreCuenta { get; init; } = string.Empty;
    public string CorreoPrincipal { get; init; } = string.Empty;
    public string? TelefonoPrincipal { get; init; }
    public string TipoPlan { get; init; } = "TRIAL";
    public string EstadoSuscripcion { get; init; } = "TRIAL";
    public bool EsPrueba { get; init; }
    public DateOnly? FechaInicioPrueba { get; init; }
    public DateOnly? FechaFinPrueba { get; init; }
    public DateOnly? FechaInicioPlan { get; init; }
    public DateOnly? FechaFinPlan { get; init; }
    public int? EmpresasPermitidas { get; init; }
    public int? UsuariosPermitidos { get; init; }
    public bool EstadoCuenta { get; init; }
    public bool EstadoEmpresa { get; init; }
    public string? Observacion { get; init; }
    public DateOnly? FechaVencimiento { get; init; }
    public int? DiasParaVencer { get; init; }
    public string ResumenVigencia { get; init; } = string.Empty;
}
