namespace SistemaControlEspaciosDeportivosWeb.Services;

public interface IDesafioEmailNotificationService
{
    Task NotifyDesafioReceivedAsync(int desafioId);
}
