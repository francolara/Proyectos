using Microsoft.Extensions.Options;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public class ReservaAutomationHostedService(
    IServiceScopeFactory scopeFactory,
    IOptions<AutomationSettings> options,
    ILogger<ReservaAutomationHostedService> logger) : BackgroundService
{
    private readonly AutomationSettings _settings = options.Value;

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                if (_settings.Enabled)
                {
                    await EjecutarCicloAsync(stoppingToken);
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error en automatizacion de reservas.");
            }

            var intervalo = Math.Max(30, _settings.IntervalSeconds);
            await Task.Delay(TimeSpan.FromSeconds(intervalo), stoppingToken);
        }
    }

    private async Task EjecutarCicloAsync(CancellationToken cancellationToken)
    {
        using var scope = scopeFactory.CreateScope();
        var spService = scope.ServiceProvider.GetRequiredService<ISportCenterStoredProcedureService>();
        var emailService = scope.ServiceProvider.GetRequiredService<INotificacionEmailService>();

        var ahora = DateTime.Now;
        var pendientes = await spService.ReservasRecordatoriosPendientesAsync(ahora);
        var enviados = 0;

        foreach (var reserva in pendientes)
        {
            if (cancellationToken.IsCancellationRequested) break;

            try
            {
                var ok = await emailService.EnviarRecordatorioReservaAsync(reserva);
                if (!ok) continue;

                await spService.ReservasMarcarRecordatorioEnviadoAsync(
                    reserva.NegocioId,
                    reserva.ReservaId,
                    _settings.UsuarioSistema);
                enviados++;
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "No se pudo enviar recordatorio de la reserva {ReservaId}.", reserva.ReservaId);
            }
        }

        var autoNoShow = await spService.ReservasAutoNoShowAsync(
            ahora,
            _settings.UsuarioSistema);

        if (enviados > 0 || autoNoShow > 0)
        {
            logger.LogInformation(
                        "Automatizacion reservas: recordatorios enviados={Enviados}, no asistio auto={AutoNoShow}.",
                enviados,
                autoNoShow);
        }
    }
}
