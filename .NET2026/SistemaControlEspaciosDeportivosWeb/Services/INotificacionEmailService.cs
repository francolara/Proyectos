using SistemaControlEspaciosDeportivosWeb.ViewModels;

namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface INotificacionEmailService
{
    Task<bool> EnviarSolicitudRecibidaAsync(SolicitudNotificacionEmailViewModel solicitud);
    Task<bool> EnviarRecordatorioReservaAsync(ReservaRecordatorioPendienteViewModel reserva);
}
