namespace SistemaAdministrativoWeb.Infrastructure.Contabilidad;

public sealed class PeriodoContableService(IPeriodoContableRepository periodoContableRepository) : IPeriodoContableService
{
    public async Task<PeriodoContableEstadoDto> ObtenerEstadoAsync(int idEmpresa, short anio, byte mes, CancellationToken cancellationToken = default)
    {
        var periodo = $"{anio:0000}{mes:00}";
        return await periodoContableRepository.ObtenerAsync(idEmpresa, periodo, cancellationToken)
            ?? new PeriodoContableEstadoDto
            {
                IdEmpresa = idEmpresa,
                Periodo = periodo,
                Cerrado = false,
                FechaRegistro = DateTime.Now
            };
    }

    public async Task<bool> EstaCerradoAsync(int idEmpresa, short anio, byte mes, CancellationToken cancellationToken = default)
    {
        var estado = await ObtenerEstadoAsync(idEmpresa, anio, mes, cancellationToken);
        return estado.Cerrado;
    }

    public Task<PeriodoContableEstadoDto> GuardarEstadoAsync(int idEmpresa, short anio, byte mes, bool cerrado, string? usuarioRegistro, CancellationToken cancellationToken = default)
    {
        return periodoContableRepository.GuardarAsync(new GuardarPeriodoContableEstadoRequest
        {
            IdEmpresa = idEmpresa,
            Periodo = $"{anio:0000}{mes:00}",
            Cerrado = cerrado,
            UsuarioRegistro = usuarioRegistro
        }, cancellationToken);
    }

    public string ConstruirMensajeBloqueo(short anio, byte mes)
    {
        return $"El periodo {mes:00}/{anio:0000} se encuentra cerrado. Reabra el periodo desde Proceso > Cerrar Periodo para continuar.";
    }
}
