using Microsoft.Extensions.Options;
using SistemaControlEspaciosDeportivosWeb.ViewModels;

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
        var emailService = scope.ServiceProvider.GetRequiredService<IEmailService>();

        var ahora = DateTime.Now;
        var pendientes = await spService.ReservasRecordatoriosPendientesAsync(ahora);
        var enviados = 0;

        foreach (var reserva in pendientes)
        {
            if (cancellationToken.IsCancellationRequested) break;

            try
            {
                await emailService.SendEmailAsync(
                    reserva.Correo,
                    reserva.Cliente,
                    $"Recordatorio de reserva - #{reserva.ReservaId}",
                    ConstruirHtmlRecordatorio(reserva));

                if (!string.IsNullOrWhiteSpace(reserva.CorreoNotificacion))
                {
                    await emailService.SendEmailAsync(
                        reserva.CorreoNotificacion,
                        reserva.Sede,
                        $"Recordatorio de reserva - #{reserva.ReservaId}",
                        ConstruirHtmlRecordatorio(reserva));
                }

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
        var autoCanceladas = await spService.ReservasAutoCancelarNoConfirmadasAsync(
            ahora,
            _settings.UsuarioSistema);

        if (enviados > 0 || autoNoShow > 0 || autoCanceladas > 0)
        {
            logger.LogInformation(
                        "Automatizacion reservas: recordatorios enviados={Enviados}, no asistio auto={AutoNoShow}, canceladas no confirmadas={AutoCanceladas}.",
                enviados,
                autoNoShow,
                autoCanceladas);
        }
    }

    private static string ConstruirHtmlRecordatorio(ReservaRecordatorioPendienteViewModel reserva)
    {
        var whatsapp = string.IsNullOrWhiteSpace(reserva.WhatsappContacto)
            ? string.Empty
            : $"<p><strong>WhatsApp de contacto:</strong> {reserva.WhatsappContacto}</p>";

        return
$"""
<h2>La Zona Deportiva</h2>
<p>Hola {reserva.Cliente},</p>
<p>Te recordamos tu reserva programada.</p>
<ul>
  <li><strong>Reserva:</strong> #{reserva.ReservaId}</li>
  <li><strong>Sede:</strong> {reserva.Sede}</li>
  <li><strong>Espacio:</strong> {reserva.Espacio}</li>
  <li><strong>Fecha:</strong> {reserva.Fecha:dd/MM/yyyy}</li>
  <li><strong>Horario:</strong> {reserva.HoraInicio:HH\:mm} - {reserva.HoraFin:HH\:mm}</li>
</ul>
{whatsapp}
<p>Gracias por usar La Zona Deportiva.</p>
<p><small>Este correo fue enviado automaticamente.</small></p>
""";
    }
}
