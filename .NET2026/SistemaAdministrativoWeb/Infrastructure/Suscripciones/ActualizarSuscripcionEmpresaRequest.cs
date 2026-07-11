namespace SistemaAdministrativoWeb.Infrastructure.Suscripciones;

public sealed class ActualizarSuscripcionCuentaRequest
{
    public int IdCuentaAdministradora { get; init; }
    public string TipoPlan { get; init; } = "TRIAL";
    public string EstadoSuscripcion { get; init; } = "TRIAL";
    public bool EsPrueba { get; init; }
    public DateOnly? FechaInicioPrueba { get; init; }
    public DateOnly? FechaFinPrueba { get; init; }
    public DateOnly? FechaInicioPlan { get; init; }
    public DateOnly? FechaFinPlan { get; init; }
    public string? TipoCobro { get; init; }
    public int DiasGracia { get; init; } = 5;
    public int? EmpresasPermitidas { get; init; }
    public int? UsuariosPermitidos { get; init; }
    public bool Activo { get; init; }
    public bool EstadoCuenta { get; init; }
    public string? Observacion { get; init; }
    public string? UsuarioRegistro { get; init; }
}
